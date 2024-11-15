import psutil
import subprocess
import platform
import socket


def get_connection_data():
    net_if_addrs = psutil.net_if_addrs()  # Get network interfaces and their addresses
    net_if_stats = psutil.net_if_stats()  # Get network interfaces' stats (up/down status)
    connection_data = {}

    # Iterate over network interfaces
    for interface in net_if_addrs:
        if interface in ['Ethernet', 'Wi-Fi', 'Ethernet 2']:  # You can adjust this list if needed
            status = net_if_stats[interface].isup  # Check if the interface is UP
            status_str = 'UP' if status else 'DOWN'

            if status_str == 'UP':
                # Initialize the interface data dictionary
                interface_data = {'status': status_str}

                # Fetch IP address (for IPv4 only)
                ip_address = None
                for addr in net_if_addrs[interface]:
                    if addr.family == socket.AF_INET:  # Use socket.AF_INET for IPv4
                        ip_address = addr.address
                        break  # We only care about the first found IP address

                if ip_address:
                    interface_data['IP'] = ip_address
                else:
                    interface_data['IP'] = 'N/A'  # No IP address found

                # Fetch additional info for Wi-Fi interfaces (SSID, security, channel, link speed, signal strength, etc.)
                if interface.startswith('Wi-Fi'):
                    if platform.system() == 'Windows':
                        # On Windows, use netsh to get Wi-Fi interface details
                        cmd = 'netsh wlan show interfaces'
                        result = subprocess.run(cmd, capture_output=True, text=True, shell=True)

                        if result.returncode == 0:
                            print("Raw output from 'netsh wlan show interfaces':")
                            print(result.stdout)  # Print the raw output for debugging

                            details = result.stdout.split('\n')
                            ssid_found = False  # Flag to check if SSID is found

                            # Extract information from netsh output
                            for line in details:
                                # Extract SSID (Network Name)
                                if 'SSID' in line:
                                    ssid = line.split(':')[1].strip()
                                    if ssid:
                                        interface_data['Network Name'] = ssid
                                        ssid_found = True

                                # Extract Radio Type (Protocol)
                                if 'Radio type' in line:
                                    protocol = line.split(':')[1].strip()
                                    if protocol:
                                        interface_data['Protocol'] = protocol

                                # Extract Security (Authentication)
                                if 'Authentication' in line:
                                    security = line.split(':')[1].strip()
                                    if security:
                                        interface_data['Security'] = security

                                # Extract Channel
                                if 'Channel' in line:
                                    channel = line.split(':')[1].strip()
                                    if channel:
                                        interface_data['Channel'] = channel

                                # Extract Link Speed (Receive rate)
                                if 'Receive rate' in line:
                                    link_speed = line.split(':')[1].strip()
                                    if link_speed:
                                        interface_data['Link Speed'] = link_speed

                                # Extract Signal Strength
                                if 'Signal' in line:
                                    signal_strength = line.split(':')[1].strip()
                                    if signal_strength:
                                        interface_data['Signal Strength'] = signal_strength

                            # If no SSID found in the output, set a fallback
                            if not ssid_found:
                                interface_data['Network Name'] = 'N/A'

                    else:
                        interface_data['Network Name'] = 'Not available on non-Windows OS'

                # Add the collected data for the interface to the dictionary
                connection_data[interface] = interface_data

    print("CONNECTION DATA", connection_data)
    return connection_data


# Example usage
connection_details = get_connection_data()
