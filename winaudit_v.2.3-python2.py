import os, datetime, sys, re
import xml.etree.ElementTree as ET

def main(argv):
    # Welcome banner
    print '*==============================================================================*'
    print '* WinAudit.py - v2.3                                                           *'
    print '* For the most up to date version visit https://github.com/scott-be/winaudit   *'
    print '* Happy Hacking!                                                               *'
    print '*==============================================================================*'

    # Get the file path if no file path is provided as a argument
    folderpath = raw_input('Enter a file path: ').strip().replace('"','') if len(argv) == 1 else argv[1]
    
    # Output network info?
    direction = ''
    while direction.strip().lower() != 'n' and direction.strip().lower() != 'y':
        direction = raw_input('Output network information [y/N]')
        if not direction:
            break

    if not direction or direction.strip().lower() == 'n':
        output_general_info(folderpath)
    else:
        output_network_info(folderpath)
    
def output_general_info(folderpath):
    # Create output folder if it does not exists...
    if not os.path.exists('output'):
        os.makedirs('output') # Create an output dir

    # Saves the file as the current datetime + _output.txt in the output dir
    time_now        = datetime.datetime.now().strftime("%Y-%m-%d-%I.%M.%S")
    output_filename = 'output' + os.sep + time_now + '_output.txt' 
    output_file     = open(output_filename,'w')

    # Heading
    output_file.write('Location\t'+
                      'Date of WinAudit\t'+
                      'Computer Name\t'+
                      # 'Computer Type\t'+
                      'AutoLogon Enabled\t'+
                      'Screen Saver Enabled\t'+
                      'Screen Saver Timeout\t'+
                      'Screen Saver Password Protected\t'+
                      'Force Network Logoff\t'+
                      'Minimum Password Length\t'+
                      'Maximum Password Age\t'+
                      'Historical Passwords\t'+
                      'Lockout Threshold\t'+
                      'Encryption\t'+
                      'Unique User ID\t'+
                      'Last Installed Security Update\n')

    # Recursivly find all .xml files in the file path
    winaudit_files = [os.path.join(dirpath, f)
        for dirpath, dirnames, files in os.walk(folderpath)
        for f in files if f.endswith('.xml')]

    # Loop through the files
    num_files_scanned = 0 # Variable to keep track of files scanned
    error_files = []
    for filename in winaudit_files:
        while True:
            try:
                    line  = ''

                # Make a tree from the xml file
                    tree = ET.parse(filename)

                # Pull - Computer Name
                    computer_name = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[1]/fieldvalue[2]").text
                    computer_name = computer_name.upper()
                    
                # Pull - Location (takes the filename and removes any slashes, the computer name and the .xml extension)
                    location = re.sub(r'^(.*[\\\/])', '', filename).replace(computer_name,'').replace(computer_name.lower(),'').replace('.xml','')
                    location = re.sub(r'^ * - *', '', location).replace('-','')
                    location = re.sub(r'^_', '', location)
                    location = location.strip()
                    if not location:
                        location = '-'

                # Pull - Date Scanned from the xml file (via RegEx)
                    title  = tree.find('./title').text
                    m  = re.search('(\d{1,2}/\d{1,2}/\d{4})', title)
                    if m:
                        date_created = m.group(1)
                    else:
                        date_created = 'Unknown'
                    
                # Add - Location, Date Scanned, Computer Name
                    line += location + ' \t'
                    line += date_created + '\t'
                    line += computer_name + '\t'

                # # Determine the computer type
                #     if computer_name[0].upper() == 'W':
                #         computer_type = 'Workstation'
                #     elif computer_name[0].upper() == 'L':
                #         computer_type = 'Laptop'
                #     else:
                #         computer_type = 'Unknown'
                #     line += computer_type + '\t'

                # Pull & Add - Security Settings
                    for child in tree.findall("./category[@title='Security']/subcategory[@title='Security Settings']/recordset/datarow"):
                        if child[0].text in {"AutoLogon", "Screen Saver", "All Accounts"}:
                            line += child[2].text + '\t'

                # Placeholder for encryption feild
                    line += '-\t'

                # Pull & Add Username (only if the username is different from the computer name)
                    username = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[17]/fieldvalue[2]").text
                    # username = '' if username.upper() == computer_name.upper() else username
                    line += username + '\t'

                # Get last installed update 
                    updates = tree.find("./category[@title='Installed Software']/subcategory[@title='Software Updates']/recordset")
                    line += str(getLatestUpdate(updates)) + '\n'

                # Done
                    print '[Done] -', computer_name, '(' + location + ')'

                # Write line to file
                    output_file.write(line)
                    num_files_scanned += 1
                    break
                
            except Exception as e:
                print '[Error] - "' + os.path.basename(filename) + '"'

                # If the error is because of a control character that cannot be parsed - Attempt to fix
                # TODO prevent any infinite loops
                if re.search('^not well-formed \(invalid token\): line ', str(e)):
                    print '\tAttempting to fix...'
                    linenum = re.search('^not well-formed \(invalid token\): line (\d*)', str(e)).group(1)
                    remove_line(filename, linenum)
                    continue
                else:
                    error_files.append(os.path.basename(filename))
                    print e
                    break

    output_file.close()
    percent_complete = round(((float(num_files_scanned)/(len(winaudit_files)))*100.0), 2)
    print '===================' 
    print '[Complete]' 
    print 'Output file: ' + output_filename 
    print 'Scanned ' + str(num_files_scanned) + ' out of ' + str(len(winaudit_files)) + ' WinAudit files. (' + str(percent_complete) + '%)'
    if error_files:
        print '\n***The following files had errors and could not be recovered:***'
        for f in error_files:
            print f
    print '===================' 
    
    # Transpose file?
    transpose = ''
    while transpose.strip().lower() != 'n' and transpose.strip().lower() != 'y':
        transpose = raw_input('Transpose output? [Y/n]')
        if not transpose:
            break
    if not transpose or transpose.strip().lower() == 'y':
        transpose_file(output_filename)

    # Done and exit
    print '==================='
    raw_input('Press any key to exit')

def output_network_info(folderpath):
    print 'output network info...'
    # Create output folder if it does not exists...
    if not os.path.exists('output'):
        os.makedirs('output') # Create an output dir

    # Saves the file as the current datetime + _output.txt in the output dir
    time_now        = datetime.datetime.now().strftime("%Y-%m-%d-%I.%M.%S")
    output_filename = 'output' + os.sep + time_now + '_network_output.txt' 
    output_file     = open(output_filename,'w')

    # Heading
    output_file.write('Computer Name\t'+
                      'Computer Location\t'+
                      'Interface Name(s)\t'+
                      'IP Address\t'+
                      'DHCP Server\t'+
                      'MAC Address\n')

    # Recursivly find all .xml files in the file path
    winaudit_files = [os.path.join(dirpath, f)
        for dirpath, dirnames, files in os.walk(folderpath)
        for f in files if f.endswith('.xml')]

    # Loop through the files
    num_files_scanned = 0 # Variable to keep track of errors
    error_files = []
    for filename in winaudit_files:
        while True:
            try:
                    line  = ''

                # Make a tree from the xml file
                    tree = ET.parse(filename)

                # Pull the computer name
                    computer_name = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[1]/fieldvalue[2]").text
                    computer_name = computer_name.upper()
                    line += computer_name + '\t'

                # Pull - Location (takes the filename and removes any slashes, the computer name and the .xml extension)
                    location = re.sub(r'^(.*[\\\/])', '', filename).replace(computer_name,'').replace(computer_name.lower(),'').replace('.xml','')
                    location = re.sub(r'^ * - *', '', location).replace('-','')
                    location = re.sub(r'^_', '', location)
                    location = location.strip()
                    if not location:
                        location = '-'
                    line += location + '\t'

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

                # Done
                    print '[Done] -', computer_name

                # Write line to file
                    output_file.write(line)
                    num_files_scanned += 1
                    break

            except Exception as e:
                print '[Error] - "' + os.path.basename(filename) + '"'

                # If the error is because of a control character that cannot be parsed - Attempt to fix
                # TODO prevent any infinite loops
                if re.search('^not well-formed \(invalid token\): line ', str(e)):
                    print '\tAttempting to fix...'
                    linenum = re.search('^not well-formed \(invalid token\): line (\d*)', str(e)).group(1)
                    remove_line(filename, linenum)
                    continue
                else:
                    error_files.append(os.path.basename(filename))
                    print e
                    break

    # Done and exit
    output_file.close()
    percent_complete = round(((float(num_files_scanned)/(len(winaudit_files)))*100.0), 2)
    print '===================' 
    print '[Complete]' 
    print 'Output file: ' + output_filename 
    print 'Scanned ' + str(num_files_scanned) + ' out of ' + str(len(winaudit_files)) + ' WinAudit files. Thats ' + str(percent_complete) + '%!'
    if error_files:
        print '\n***The following files had errors and could not be recovered:***'
        for f in error_files:
            print f
    print '==================='
    raw_input('Press any key to exit')

def getLatestUpdate(tree):
  date_list = [] # Used to store list of dates
  for recordset in tree:
    if str(recordset.tag) == "datarow":
      update_description = recordset.find("fieldvalue[3]").text # Store update description
      update_date = recordset.find("fieldvalue[2]").text # Store installed on date
      
      # Match only security updates descriptions
      if re.match(r'(Security Update$)|(Security Update for Windows XP(.*))|(Security Update for Windows 7(.*))', str(update_description), re.IGNORECASE):
        if update_date is not None:
            if '-' in update_date:
                date_list.append(datetime.datetime.strptime(update_date, '%Y-%m-').date())
            elif '/' in update_date:
                date_list.append(datetime.datetime.strptime(update_date, '%m/%d/%Y').date())
            else:
                return 'error'

  # Sort list of dates
  date_list = sorted(date_list)

  # return latest date (last element)
  return str(date_list[-1])

def transpose_file(output_filename):
    print 'Transposing file...',
    
    with open(output_filename, 'r') as f:
        lis = [x.strip().split('\t') for x in f]

    output_file = open(output_filename, 'w') # Have to open the file back up
    line = ''
    for x in zip(*lis):
        for y in x:
            line += y + '\t'
        line += '\n'
    output_file.write(line)
    output_file.close
    print '[Done]'

def remove_line(file_name, line_num):
    line_num = int(line_num) - 1
    lines = open(file_name, 'r').readlines()
    lines[line_num] = '\n'
    out = open(file_name, 'w')
    out.writelines(lines)
    out.close()

if __name__ == "__main__":
    main(sys.argv) 
