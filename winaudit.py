import os, time, sys, re
import xml.etree.ElementTree as ET

def main(argv):
    # Get the file path if no file path is provided
    folderpath = raw_input('Enter a file path: ').strip().replace('\\', '') if len(argv) == 1 else argv[1]

    # Open output file
    if not os.path.exists('output'):
        os.makedirs('output') # Create an output dir
    epoch           = str(int(time.time()))
    output_filename = 'output' + os.sep + epoch + '_output.txt' # saves the file as the current epoch time + _output.txt in the output dir
    output_file     = open(output_filename,'w')

    # Recursivly find all .xml files in the path
    winaudit_files = [os.path.join(dirpath, f)
        for dirpath, dirnames, files in os.walk(folderpath)
        for f in files if f.endswith('.xml')]

    # Loop through the files
    num_errors = 0
    for filename in winaudit_files:
        try:
            tree          = ET.parse(filename) # make a tree from the xml file
            date_created  = time.strftime('%m-%d-%Y', time.localtime(os.path.getctime(filename))) # Print the date the file was created
            computer_name = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[1]/fieldvalue[2]").text
            computer_type = "Workstation" if computer_name[0] == 'W' else 'Laptop'
            username      = tree.find("./category[@title='System Overview']/subcategory/recordset/datarow[17]/fieldvalue[2]").text
            username      = '' if username.upper() == computer_name.upper() else username
            location      = re.sub(r'^(.*[\\\/])','', filename).replace("_" + computer_name + ".xml", "").replace('_', ' - ')
            computer_name = computer_name.upper() 
            line          = '' # Line used to write to file

            print date_created + '\t',
            print location + '\t',
            print computer_name + '\t',
            print computer_type + '\t',
            
            line += date_created + '\t'
            line += location + '\t'
            line += computer_name + '\t'
            line += computer_type + '\t'

            # Print security settings
            for child in tree.findall("./category[@title='Security']/subcategory[@title='Security Settings']/recordset/datarow"):
                if child[0].text in {"AutoLogon", "Screen Saver", "All Accounts"}:
                    print child[2].text + '\t',
                    line += child[2].text + '\t'
            print '\t', # Skip the encryption column
            print username

            line += '\t'
            line += username + '\n'

            # Write line to file
            output_file.write(line)

        except Exception, e:
            print "ERROR: Can't read from:", filename
            error_file = open ('output' + os.sep + epoch + '_errors.txt', 'a')
            error_file.write(filename + '\n')
            error_file.close()
            num_errors += 1

    output_file.close()
    print '==================='
    print '[done]'
    print 'Output file:' + output_filename
    print 'Scanned ' + str(len(winaudit_files)) + ' files with ' + str(num_errors) + ' errors'

if __name__ == "__main__":
    main(sys.argv) 