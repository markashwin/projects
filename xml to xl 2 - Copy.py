import xml.etree.ElementTree as ET
import pandas as pd
import xlsxwriter
import glob


df_list = list()
files = glob.glob(r'C:/Users/Mark Ashwin/Documents/TCPL Managed VA/Web_application_XML_file/*.xml', 
                   recursive = True)

for index,file in enumerate(files):
        print(file)
        tree = ET.parse(file)
        root = tree.getroot()
        last_seen = root.find('scan-information').find('scan-date-and-time').text.split()[0]
        origin = root.find('scan-configuration').find('starting-url').text

        for child in root:
                if child.tag == "issue-group":
                        ig = child
                
                if child.tag == "issue-type-group":
                        itg = child

                if child.tag == "fix-recommendation-group":
                        frg = child

                if child.tag == "url-group":
                        ug = child
                
                if child.tag == "security-risk-group":
                        srg = child

                if child.tag == "advisory-group":
                        ag = child


        # 0 - Issue Name
        # 1 - Description
        # 2 - Origin
        # 3 - Affected URLs
        # 4 - Severity
        # 5 - Last Seen
        # 6 - Remediation
        

        issue_name_list = list()
        description_list = list()
        origin_list = list()
        affected_urls_list = list()
        severity_list = list()
        last_seen_list = list()
        remediation_list = list()


        for child in itg:

                issue_name_list.append(child[0].text) # Adding name of the issue

                description = '' 
                risk = ''
                for ref in child.find('security-risks').findall('ref'): #ref.text is the reference id of risk
                        for elem in srg:
                                if ref.text == elem.get('id'):
                                        risk = risk + elem.text + '\n'
                reasoning = '' #getting the reasoning text of the issue
                for child1 in ig:
                        if child.get('id') == child1.find('issue-type')[0].text:    
                                vg = child1.find('variant-group')[0] 
                                #print(child.get('id'),'||',vg.find('reasoning').text)   
                                if vg.find('reasoning').text:                                        
                                        reasoning = vg.find('reasoning').text + '\n'
                                #print(child.get('id'),'||',reasoning)
                                try:
                                        if vg.find('issue-information').find('template').text == 'CipherSuites_Template':
                                                for tip in vg.find('issue-information').find('issue-tips').findall('issue-tip'):
                                                        reasoning = reasoning + tip.text + '\n'
                                                        reasoning = reasoning +'\n'
                                                i=1
                                                for ciphersuite in vg.find('issue-information').find('CipherSuites').findall('CipherSuite'):
                                                        reasoning = reasoning + str(i) + '. ' + ciphersuite.find('Name').text + '\n'
                                                        i = i + 1
                                                break
                                except AttributeError:
                                        pass
                description = risk + reasoning
                description_list.append(description) # Adding description of the issue      

                origin_list.append(origin) # Adding origin of the issue

                url_list = set() 
                for child1 in ig: 
                        if child1.find('issue-type')[0].text == child.get('id'):
                                url_id = child1.find('url')[0].text
                                for url in ug:
                                        if url_id == url.get('id'):
                                                url_list.add(url[0].text.replace(origin,''))
                affected_urls_list.append('\n'.join(url_list)) # Adding path of the issue
                
                for child1 in ig:  
                        if child.get('id') == child1.find('issue-type')[0].text:
                                severity = child1.find('severity').text
                                severity_list.append(severity)
                                break

                last_seen_list.append(last_seen) # Adding last seen of the issue
                        
                for child1 in frg: 
                        if child.find('fix-recommendation')[0].text == child1.get('id'): # Adding recommendation text to str1
                                str1 = '' 
                                for element in child1[0][0]:#.findall():                                
                                        try:
                                                str1 = str1 + element.text + '\n'
                                        except:
                                                pass                        
                                remediation_list.append(str1) # Adding remediation of the issue
                        
        data = {
                
                'Issue Name':issue_name_list,
                'Description':description_list,
                'Origin':origin_list,
                'Path':affected_urls_list,
                'Severity':severity_list,
                'Last Seen':last_seen_list,
                'Remediation':remediation_list
        }
        #print('hi')
        #print(len(data['Origin']),len(data['Path']),len(data['Severity']),len(data['Last Seen']),len(data['Remediation']))
        
        df = pd.DataFrame(data)
        df_list.append(df)
        rows = len(df.index)
        columns = len(df.columns)
        print(f'origin = {origin}, no of issues = {rows}')

        if index == len(files)-1:
                print(f'\nAll {len(files)} files done ...\n')

df = pd.concat(df_list)

severity_order = ['Critical','High','Medium','Low','Informational']
df['Severity'] = pd.Categorical(df['Severity'], categories=severity_order, ordered=True)
df = df.sort_values(by='Severity')

writer = pd.ExcelWriter(r'C:\Users\Mark Ashwin\Documents\TCPL Managed VA\Web_application_XML_file\output.xlsx', engine="xlsxwriter", engine_kwargs={'options': {'strings_to_urls': False}})

df.to_excel(writer, sheet_name='Issue List', index=False)
wb = writer.book
ws = writer.sheets["Issue List"]

table_format = wb.add_format()
table_format.set_border()

header_row_format = wb.add_format()
header_row_format.set_bold()
header_row_format.set_bg_color('#B4C6E7')
header_row_format.set_size(11)
header_row_format.set_font_name('Arial')
header_row_format.set_border()

medium_col_format = wb.add_format()
medium_col_format.set_bold()
medium_col_format.set_bg_color('#FFC000')
medium_col_format.set_size(11)
medium_col_format.set_font_name('Arial')
medium_col_format.set_align('center')
medium_col_format.set_align('vcenter')

high_col_format = wb.add_format()
high_col_format.set_bold()
high_col_format.set_bg_color('#FF0000')
high_col_format.set_size(11)
high_col_format.set_font_name('Arial')
high_col_format.set_align('center')
high_col_format.set_align('vcenter')

low_col_format = wb.add_format()
low_col_format.set_bold()
low_col_format.set_bg_color('#92D050')
low_col_format.set_size(11)
low_col_format.set_font_name('Arial')
low_col_format.set_align('center')
low_col_format.set_align('vcenter')

informational_col_format = wb.add_format()
informational_col_format.set_bold()
informational_col_format.set_bg_color('#00B0F0')
informational_col_format.set_size(11)
informational_col_format.set_font_name('Arial')
informational_col_format.set_align('center')
informational_col_format.set_align('vcenter')

lastseen_format = wb.add_format({
                                        'font_name':'Arial',
                                        'font_size':11,
                                        'align':'center',
                                        'valign':'vcenter',
                                })

origin_format = wb.add_format({
                                        'font_name':'Arial',
                                        'font_size':11,
                                        'align':'left',
                                        'valign':'vcenter',
                                })


name_format = wb.add_format()
name_format.set_font_name('Arial')
name_format.set_size(11)
name_format.set_bold()
name_format.set_align('center')
name_format.set_align('vcenter')

des_format = wb.add_format()
des_format.set_font_name('Arial')
des_format.set_size(11)
des_format.set_align('left')
des_format.set_align('vcenter')

url_col_format = wb.add_format()
url_col_format.set_font_name('Arial')
url_col_format.set_size(11)
url_col_format.set_align('left')
url_col_format.set_align('vcenter')

sev_format = wb.add_format()
sev_format.set_font_name('Arial')
sev_format.set_size(11)
sev_format.set_bold()
sev_format.set_align('center')
sev_format.set_align('vcenter')

rem_format = wb.add_format()
rem_format.set_font_name('Arial')
rem_format.set_size(11)
rem_format.set_align('left')
rem_format.set_align('vcenter')


rows = len(df.index)
columns = len(df.columns)
print(f'rows = {rows}, columns = {columns}')
colname = xlsxwriter.utility.xl_col_to_name(columns-1)
  
for i in range(rows): 
  
    # Rewriting the urls as string
    ws.write_string("C"+str(i+2),df.iloc[i,2])

# Applying Formatting for header row
ws.conditional_format('A1:' + colname + '1', {'type': 'cell',
                                     'criteria': '>=',
                                     'value': 0, 'format':header_row_format })


# Applying formatting for whole table
ws.conditional_format('A1:' + colname + str(rows+1), {'type': 'cell',
                                     'criteria': '>=',
                                     'value': 0, 'format':table_format })

# Applying formatting for medium issues
ws.conditional_format('E2:E' + str(rows+1) , {'type': 'text',
                                     'criteria':'containing',
                                     'value':'Medium', 'format':medium_col_format })

# Applying formatting for high issues
ws.conditional_format('E2:E' + str(rows+1) , {'type': 'text',
                                     'criteria':'containing',
                                     'value':'High', 'format':high_col_format })

# Applying formatting for low issues
ws.conditional_format('E2:E' + str(rows+1) , {'type': 'text',
                                     'criteria':'containing',
                                     'value':'Low', 'format':low_col_format })

# Applying formatting for informational issues
ws.conditional_format('E2:E' + str(rows+1) , {'type': 'text',
                                     'criteria':'containing',
                                     'value':'Informational', 'format':informational_col_format })


ws.set_column(0, 0, 46.5, name_format)
ws.set_column(1, 1, 42, des_format)
ws.set_column(2, 2, 36, origin_format)
ws.set_column(3, 3, 32, url_col_format)
ws.set_column(4, 4, 16.5, sev_format)
ws.set_column(5, 5, 15, lastseen_format)
ws.set_column(6, 6, 81, rem_format)

writer.close()   


     
        





