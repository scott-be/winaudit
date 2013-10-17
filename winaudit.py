import os, time, sys, re
import xml.etree.ElementTree as ET

def main(argv):
    # Get the file path if no file path is provided as a argument
    folderpath = raw_input('Enter a file path: ').strip().replace('"','') if len(argv) == 1 else argv[1]

    # Create output folder if it does not exists...
    if not os.path.exists('output'):
        os.makedirs('output') # Create an output dir

    # Saves the file as the current epoch time + _output.txt in the output dir
    epoch           = str(int(time.time()))
    output_filename = 'output' + os.sep + epoch + '_output.txt' 
    output_file     = open(output_filename,'w')
    # Heading
    output_file.write('Date Scanned\t'+
                        'Location\t'+
                        'Computer Name\t'+
                        'Computer Type\t'+
                        'AutoLogon Enabled\t'+
                        'Screen Saver Enabled\t'+
                        'Screen Saver Timeout\t'+
                        'Screen Saver Password Protected\t'+
                        'Force Network Logoff\t'+
                        'Minimum Password Length\t'+
                        'Maximum Password Age\t'+
                        'Historical Passwords\t'+
                        'Lockout Threshold\t'+
                        '[skip]\t'+
                        'Username\t'+
                        'Interface Name(s)\t'+
                        'IP Address\t'+
                        'DHCP Server\t'+
                        'MAC Address\n')

    # Recursivly find all .xml files in the file path
    winaudit_files = [os.path.join(dirpath, f)
        for dirpath, dirnames, files in os.walk(folderpath)
        for f in files if f.endswith('.xml')]

    # Loop through the files
    num_errors = 0 # Variable to keep track of errors
    for filename in winaudit_files:
        try:
                line  = ''

            # Make a tree from the xml file
                tree = ET.parse(filename)

            # Pull the date of scan from the xml file (via RegEx)
                title  = tree.find('./title').text
                m  = re.search('(\d{1,2}/\d{1,2}/\d{4})', title)
                if m:
                    date_created = m.group(1)
                else:
                    date_created = 'Unknown'
                line += date_created + '\t'
            # Pull the computer name
                computer_name = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[1]/fieldvalue[2]").text
                computer_name = computer_name.upper()

            # Pull the location (takes the filename and removes any slashes, the computer name and the .xml extension)
                location = re.sub(r'^(.*[\\\/])', '', filename).replace(computer_name,'').replace(computer_name.lower(),'').replace('.xml','')
                line    += location + '\t'

            # Add computer name
                line += computer_name + '\t'

            # Determine the computer type
                if computer_name[0].upper() == 'W':
                    computer_type = 'Workstation'
                elif computer_name[0].upper() == 'L':
                    computer_type = 'Laptop'
                else:
                    computer_type = 'Unknown'
                line += computer_type + '\t'

            # Print security settings
                for child in tree.findall("./category[@title='Security']/subcategory[@title='Security Settings']/recordset/datarow"):
                    if child[0].text in {"AutoLogon", "Screen Saver", "All Accounts"}:
                        line += child[2].text + '\t'

                line += '\t' # Skip column

            #Pull the username (only if the username is different from the computer name)
                username = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[17]/fieldvalue[2]").text
                username = '' if username.upper() == computer_name.upper() else username    
                line += username + '\t'

            # Pull IP Address, MAC address, DHCP Server from all network interfaces

                interfaces = tree.find("./category[@title='Network TCP/IP']").getchildren()
                for i, interface in enumerate(interfaces):

                    interface_name = interface.get('title')
                    ip_address = interface.find('recordset/datarow[10]/fieldvalue[2]').text
                    dhcp_server = interface.find('recordset/datarow[9]/fieldvalue[2]').text
                    mac_address = interface.find('recordset/datarow[16]/fieldvalue[2]').text

                    if ip_address == None:   # look to see if the IP address was found
                        ip_address = "None"
                    if dhcp_server == None:  # look to see if DHCP Server was found
                        dhcp_server = "None"
                    if mac_address == None:  # look to see if a MAC was found
                        mac_address = "None"

                    line += interface_name + '\t'
                    line += ip_address + '\t'
                    line += dhcp_server + '\t'
                    line += mac_address + '\t'

                    if i == len(interfaces)-1: # Look to see if its the last interface and insert a newline
                        line += '\n'

                print '[Done] - ', computer_name

            # Write line to file
                output_file.write(line)

        except IOError, e:
            print "ERROR: Can't read from:", os.path.basename(filename)
            error_file = open('output' + os.sep + epoch + '_errors.txt', 'a')
            error_file.write(filename + '\n')
            error_file.close()
            num_errors += 1
        except Exception, e2:
            print 'Error: "',os.path.basename(filename), '"" -- ', e2
            error_file = open('output' + os.sep + epoch + '_errors.txt', 'a')
            error_file.write(filename + '\n')
            error_file.close()
            num_errors += 1

    output_file.close()
    print '==================='
    print '[Complete]'
    print 'Output file: ' + output_filename
    print 'Scanned ' + str(len(winaudit_files)) + ' files with ' + str(num_errors) + ' errors'
    raw_input('Press any key to exit')

if __name__ == "__main__":
    main(sys.argv) 
