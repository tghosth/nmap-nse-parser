import xlsxwriter

# OptionParser imports
from optparse import OptionParser
from optparse import OptionGroup

VERSION = '1.0'

# Options definition
parser = OptionParser(usage="%prog [options]\nVersion: " + VERSION)

# Options definition
mandatory_grp = OptionGroup(parser, 'Mandatory parameters')
mandatory_grp.add_option('-i', '--input', help='Nmap scan output file (".nmap" format only)', nargs=1)

output_grp = OptionGroup(parser, 'Output parameters')
output_grp.add_option('-o', '--output', help='Excel file output filename (input filename with .xlsx added at the end if not specified)', nargs=1)

parser.option_groups.extend([mandatory_grp, output_grp])

class ipItem:
    def __init__(self):
        self.nseList = {}
        self.port_section = ''
        self.ip = ''



options, arguments = parser.parse_args()

filename = ''

# Input descriptor
if options.input != None:
    filename = options.input
else:
    parser.error('Please provide an input file using "--input" !')


current_ip=''
in_port_list = False
in_nse_section = False
port_section = ''
start_section = 'Nmap scan report for '
nse_section = 'Host script results:'
curr_nse_item_header = ''
curr_nse_item_text = ''
ip_item_list = []
nse_list = {}
nse_master_list = []
curr_ip_item = None

with open(filename,'r') as fileIn:
    for fileLine in fileIn:
        if start_section in fileLine:

            if curr_ip_item is not None:
                ip_item_list.append(curr_ip_item)

            curr_ip_item = ipItem()

            current_ip = fileLine.replace(start_section, '').replace('\n','')
            curr_ip_item.ip = current_ip

        elif fileLine[:4] == 'PORT':
            in_port_list = True
            port_section = fileLine
        elif in_port_list and (not fileLine == '\n'):
            port_section = port_section + fileLine
        elif in_port_list and (fileLine == '\n'):
            in_port_list = False
            curr_ip_item.port_section = port_section
            port_section = ''
        elif nse_section in fileLine:
            in_nse_section = True
            curr_nse_item_text = ''
        elif in_nse_section:
            if ((not fileLine[2:][:1] == ' ') or fileLine[:1] == '\n'):
                if len(curr_nse_item_text) > 0:
                    nse_list[curr_nse_item_header] = curr_nse_item_text
                    if curr_nse_item_header not in nse_master_list:
                        nse_master_list.append(curr_nse_item_header)

                curr_nse_item_text = ''
                curr_nse_item_header = fileLine[2:].split(':')[0]

            if fileLine[:1] == '\n':
                in_nse_section = False
                curr_ip_item.nseList = nse_list
                nse_list = {}
            else:
                curr_nse_item_text = curr_nse_item_text + fileLine

ip_item_list.append(curr_ip_item)

out_file = filename + '.xlsx'

if options.output is not None:
    out_file = options.output

with xlsxwriter.Workbook(out_file) as workbook:
    worksheet = workbook.add_worksheet()


    worksheet.write(0, 0, 'IP')
    worksheet.write(0, 1, 'Ports')

    col_num = 2
    for nse in nse_master_list:
        worksheet.write(0, col_num, nse)
        col_num = col_num + 1

    item_num = 1

    for item_out in ip_item_list:
        worksheet.write(item_num, 0, item_out.ip)
        worksheet.write(item_num, 1, item_out.port_section)
        col_num = 2
        for nse in nse_master_list:
            if nse in item_out.nseList:
                worksheet.write(item_num, col_num, item_out.nseList[nse])
            else:
                worksheet.write(item_num, col_num, '')

            col_num = col_num + 1

        item_num = item_num + 1