import urllib.request
import urllib.response
import urllib.parse
import urllib.error
import re
import xlwt
import json
from time import sleep
import xlrd
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt.Workbook import Workbook

class GetWealth:
    def __init__(self,page):
        self.baseurl="http://www.chinawealth.com.cn/lccpAllProJzyServlet.go"
        self.header={
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Referer':'http://www.chinawealth.com.cn/zzlc/jsp/lccp.jsp'
        };
        self.data={
        'cpjglb':'',
        'cpsylx':'',
        'cpyzms':'',
        'cpfxdj':'',
        'cpqx':'',
        'cpzt':'02',
        'cpdjbm':'',
        'cpmc':'',
        'cpfxjg':'',
        'mjqsrq':'',
        'mjjsrq':'',
        'areacode':'',
        'tzzlxdm':'03',
        'pagenum':str(page),
        'orderby':'',
        'code':''
      };
    def GetCurrentPageHTML(self):
        post_data=urllib.parse.urlencode(self.data).encode('utf-8');
        try:
            html_request=urllib.request.Request(self.baseurl,headers=self.header,data=post_data);
            html_open=urllib.request.urlopen(html_request,timeout=10);
            return html_open.read().decode('utf-8');
        except urllib.error.URLError as e:
            if hasattr(e, 'code'):
                print(e.code);
            print (e.reason);
            return ''
    def SaveAsExcel(self):
        json_to_dict=json.loads(self.GetCurrentPageHTML());
        if not json_to_dict['List']:
#             print ('no')
            return False;
        else:
#             print (json_to_dict)
            #save as excel
            try:
                rb=open_workbook('chinawealth.xls',formatting_info=True)
            except Exception as e:
                f=xlwt.Workbook();
                wealthsheet=f.add_sheet("wealthsheet", cell_overwrite_ok=True)
                f.save('chinawealth.xls');
                rb=open_workbook('chinawealth.xls',formatting_info=True)
            r_table=rb.sheet_by_index(0);
            nrows=r_table.nrows;
            ncols=r_table.ncols;
            
            wb=copy(rb);
            w_table=wb.get_sheet(0);
            if nrows>0:
                for i in range(len(json_to_dict['List'])):
                    j=1
                    w_table.write(i+nrows,0,i+nrows)
                    for key in json_to_dict['List'][i]:
                        w_table.write(i+nrows,j,json_to_dict['List'][i][key])
                        j+=1;
            else:
                w_table.write(0,0,json_to_dict['Count'])
                keynum=1;
                for key in json_to_dict['List'][0]:
                    w_table.write(0,keynum,key);
                    keynum+=1;
                
                for i in range(len(json_to_dict['List'])):
                    j=1
                    w_table.write(i+1,0,i+1)
                    for key in json_to_dict['List'][i]:
                        w_table.write(i+1,j,json_to_dict['List'][i][key])
                        j+=1;
            
            wb.save('chinawealth.xls')        
                
                
            return True;
        
            
            
def main():
    i=1;
    
    while True:
        CurWealthPage=GetWealth(i);
        if CurWealthPage.SaveAsExcel():
            i+=1;
            print(i)
            sleep(30);
        else:
            break;
#     for i in range(1,3):
#         CurWealthPage=GetWealth(i);
#         CurWealthPage.SaveAsExcel();
#         print (i)

if __name__=='__main__':
    main()


        