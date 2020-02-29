import glob
import csv
import re
import datetime
import sqlite3
import os
import webbrowser
import sys
from jinja2 import Environment
pth=os.environ['HOMEDRIVE']+os.environ['HOMEPATH']+'/Documents/python/'
def imp(qryTy):
        db=sqlite3.connect(pth+"stk.db")
        cursor=db.cursor()
        if qryTy in ('3','4','T'):
                for x in glob.glob('a*.csv'):
                        yr=os.path.basename(x)[1:5]
                        with open(x, newline='',encoding='utf-8') as csvfile:
                                spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
                                r1="(.*)月(.*)日"
                                r2="(.*)/(.*)"
                                cnt=0
                                for row in list(spamreader)[1:]:
                                        cnt=cnt+1
                                        t=re.match(r1,row[0])
                                        if t:
                                                t=t.groups()
                                        else:
                                                t=re.match(r2,row[0]).groups()
                                        row[0]=yr+'-'+t[0].rjust(2,'0')+'-'+t[1].rjust(2,'0')
                                        row[1]=row[1].replace('[','')
                                        row[1]=row[1].replace(']','')
                                        row[5]=row[5].replace(',','')
                                        row[6]=row[6].replace(',','')
                                        row[7]=row[7].replace(',','')
                                        row[8]=row[8].replace(',','')
                                        row[9]=row[9].replace(',','')
                                        row[10]=row[10].replace(',','')
                                        dt=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        #print(row[0],row[1],row[3],row[5],row[6],row[7],row[8],row[9],row[10],dt,'A')
                                        if cnt==1:
                                                ym=row[0][:7]
                                                if qryTy=='T':
                                                        ym=row[0][:4]
                                                cursor.execute("delete from trans where mk='A' and dtdat like ?||'%'",(ym,))
                                        cursor.execute(
                                                "insert into trans(dtdat,stk,salty,qty,prc,amt,fee,tax,net,txdat,mk) values(?,?,?,?,?,?,?,?,?,?,?)" ,
                                                (row[0],row[1],row[3],row[5],row[6],row[7],row[8],row[9],row[10],dt,'A')
                                        )
                                print('cnt:',cnt)                                
                        os.remove(x)
        for x in glob.glob('b*.csv'):
                yr=os.path.basename(x)[1:5]
                cursor.execute("delete from trans where mk='B' and dtdat like ?||'%'",(yr,))
                with open(x, newline='',encoding='utf-8') as csvfile:
                        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
                        for row in list(spamreader)[4:-3]:
                                row[1]=row[1].replace('/','-')
                                row[3]=row[3][1:2]
                                row[4]=row[4].replace(',','')
                                row[5]=row[5].replace(',','')
                                row[9]=row[9].replace(',','')
                                row[12]=row[12].replace(',','')
                                row[14]=row[14].replace(',','')
                                dt=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                cursor.execute(
                                        "insert into trans(dtdat,stk,salty,qty,prc,amt,fee,tax,net,txdat,mk) values(?,?,?,?,?,?,?,?,?,?,?)" ,
                                        (row[1],row[2],row[3],row[4],row[5],'0',row[12],row[9],row[14],dt,'B')
                                )                              
                os.remove(x)
        db.commit()
        db.close()
	
	
def qry(whereTy,data,orderTy):
        sql="select stk,dtdat,salty,qty,prc,amt,fee,tax,net,txdat from trans "+ \
                "where " +whereTy + " like ? "+ \
                "order by " + \
                orderTy
        db=sqlite3.connect(pth+"stk.db")
        cursor=db.cursor()
        cursor.execute(sql , ('%'+data+'%',))
        result = cursor.fetchall()
        db.commit()
        db.close()
        template = """
<table border="1">
{% for row in rows %}
  <tr>
  {%- for col in row -%}
	{%- if loop.index == 3 -%} 
    <td class='salTy'>{{ col }}</td>
	{%- else -%}
    <td>{{ col }}</td>
	{%- endif -%}
  {%- endfor -%}
  </tr>
{% endfor %}
</table>
"""
        return Environment().from_string(template).render(rows=result)	
if __name__ == '__main__':
        #print(sys.argv,len(sys.argv))
        if len(sys.argv)==2:
                imp(sys.argv[1])
        else:
                imp('')
        webbrowser.open("shortcuts:")


