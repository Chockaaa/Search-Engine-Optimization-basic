class density:
    def __init__(self):#Constructor to introduce the purpose of the program
        print("We are going to calulate density today\n")



    def readexcel(self):#Extract the URL link(assign to self) and set of words(return to search)
        import xlrd
        home=("seolink.xlsx")
        wb=xlrd.open_workbook(home)
        sheet=wb.sheet_by_index(0)
        
        sheet.cell_value(0,0)
        self.urllink=sheet.cell_value(0, 0)
        print("Your URL is:",self.urllink)
        search=[]
        for row in range(1,sheet.nrows):
            temp=sheet.cell_value(row,0)
            search.append(temp.lower()) 

        print("The words to calculate density:",search)
        return search
    
        
    def extract(self):#Function to extract the text from the URL link and return the text
        import requests
        import bs4
        res=requests.get(self.urllink)
        soup=bs4.BeautifulSoup(res.text,'html5lib')
        for x in soup("script","style","a"):
            x.extract()
                
        spam=soup.get_text()

        wordbank=spam
        
        wordbank=wordbank.lower()
        print(wordbank)
        return wordbank
    
    def getandarrange(self,wordbank):#Function to split the text and return the words
        import re
        pattern=re.compile("\w+")
        matchobject=pattern.findall(wordbank)        
        wordbank=matchobject  
        print(wordbank)        
        return wordbank


    def dispdensity(self,wordbank,search):#Function to calculate density of the set of words and save to database       
        totalcount=len(wordbank)
        print(totalcount)
        
        import sqlite3
        server=sqlite3.connect(':memory:')
        query="create table SEO(keyword,wordcount,density)"
        print("Table created in database")
        server.execute(query)        
        for x in range(len(search)):
            searchcount=0
            searchcount=wordbank.count(search[x])
            print(searchcount)
            print(totalcount)
            density=1
            density=(searchcount/totalcount)
            print("The density of",search[x],"is", density,"%")            
            insert="insert into SEO values(?,?,?)"
            server.execute(insert,(search[x],searchcount,density))
            server.commit()

        print("Table completed in database")
        
        selectquery="select * from SEO"
        retrieve=server.execute(selectquery)
        for x in retrieve:
            print(x[0],x[1],x[2])
            
        import xlsxwriter
        workbook=xlsxwriter.Workbook("density.xlsx")
        worksheet=workbook.add_worksheet()        
        bold=workbook.add_format({'bold':1})
        percent=workbook.add_format({'num_format':'0.0%'})
        
        worksheet.write('A1','Keyword',bold)
        worksheet.write('B1','Wordcount',bold)
        worksheet.write('C1','Density',bold)
        
        row=1
        col=0
        
        selectquery="select * from SEO"
        retrieve=server.execute(selectquery)
        for x in retrieve:
            print(x[0],x[1],x[2])
            worksheet.write_string(row,col,x[0])
            worksheet.write_number(row,col+1,x[1])
            worksheet.write_number(row,col+2,x[2],percent)
            row+=1



        worksheet2=workbook.add_worksheet()
        column_chart=workbook.add_chart({'type':'column'})

        column_chart.add_series({
            'name':         '=Sheet1!B1',
            'categories':   '=Sheet1!A2:A6',
            'values':       '=Sheet1!B2:B6',

            })
        line_chart=workbook.add_chart({'type':'line'})

        line_chart.add_series({
            'name':         '=Sheet1!C1',
            'categories':   '=Sheet1!A2:A7',
            'values':       '=Sheet1!C2:C6',
            'y2_axis':      True,
            })

        column_chart.combine(line_chart)

        column_chart.set_title({'name':'Density of Keywords'})
        column_chart.set_x_axis({'name':'Key Words'})
        column_chart.set_y_axis({'name':'Word Count'})
        column_chart.set_y2_axis({'name':'Density'})

        worksheet2.insert_chart('E2', column_chart)
        
        
        workbook.close()
   
        
    def __del__(self):#Destructor to terminate program
        print("End of Calculation")
    

try:
    SEO=density()
    search=SEO.readexcel()
    wordbank=SEO.extract()
except Exception as e:
    print(e,"Exception Caught")
else:
    wordbank=SEO.getandarrange(wordbank)
    SEO.dispdensity(wordbank,search)
finally:
    del SEO

