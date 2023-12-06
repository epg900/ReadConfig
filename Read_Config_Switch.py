from netmiko import ConnectHandler
import getpass,tqdm,os,re,json
from openpyxl import load_workbook
from docxtpl import DocxTemplate, InlineImage, RichText

'''
Word Template:
{%tr for k,v in dict1.items() %}
{{ v.ip }}	{{ v.model }}	{{ v.sn }}	{{ v.rom }}   {{ v.link }}
{%tr endfor %}

Create Zip File:
#from shutil import make_archive
#make_archive("RO_SW_Config","zip")
'''

conf={'device_type':'cisco_ios','host':'','username':'','password':''}

data="{}"

try:
    with open('ipdict.txt') as f:
        data=f.read()
except:
    pass
    
ip_dict=json.loads(data)

conf['username']=input('Enter Your Username: ')
conf['password']=getpass.getpass(prompt='Enter Your Password: ')


doc = DocxTemplate("templ.docx")
dic={}

wb=load_workbook("templ.xlsx")
ws=wb.active
sh=wb['Sheet']

try:
    for i,key in enumerate(tqdm.tqdm(ip_dict)):
        conf['host']=ip_dict[key]    
        net_obj=ConnectHandler(**conf)
        out=net_obj.send_command("show inventory")
        model=re.findall(' DESCR: "(.*)"',out)
        sn=re.findall(' SN: (.*)\n',out)
        out=net_obj.send_command("show ver")
        rom=re.split('\n',out)    
        output=net_obj.send_command("show run")
        filename=key + ".txt"
        f=open(filename,'w')
        f.write(output)
        f.close()
        
        rt=RichText()
        rt.add(key,url_id=doc.build_url_id(filename))
        dic[key]={}
        dic[key]['ip']=ip_dict[key]
        dic[key]['model']=model[0]
        dic[key]['sn']=sn[0]
        dic[key]['rom']=rom[0]
        dic[key]['link']=rt
        
        sh['A{}'.format(i+3)]= ip_dict[key]
        sh['B{}'.format(i+3)]= model[0]
        sh['C{}'.format(i+3)]= sn[0]
        sh['D{}'.format(i+3)]= rom[0]
        sh['E{}'.format(i+3)]= '=HYPERLINK("{}","{}")'.format(filename,key)
               
        net_obj.disconnect()
except:
    pass

ctx={'dict1': dic}
doc.render(ctx)
doc.save("R_SW_Config.docx")

wb.save("R_SW_Config.xlsx")

print("\nCompleted!")
os.system('pause')
