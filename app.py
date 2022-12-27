import openpyxl as opxl
import pandas as pd
import sys, os
import unicodedata as ud
import re
from random import randint
from pprint import pprint as pp
from fuzzywuzzy import fuzz
from collections import Counter
import win32ui
from pprint import pprint
from time import perf_counter

from product_class import Product

class ProductTitle(Product):

    regex_patterns = {
        'parenth': re.compile("\((.+?)\)"),
        'brand': '', #uses fuzzy search
        'series': '', #simple conditionals
        'grouping': [re.compile("([Ss])(/\d*)\s"), re.compile("(\s\d\d?)?(\s[SΣsσ][ΕεEe][TtΤτ]\s\d{1,4}\s\S+)\s")], #
        'SKU': '', #conditionals
        'volume': re.compile("\| \d+,?\d?\s?[MmLlCc][TtLlCc\s]"),
        'dimension': re.compile("\| [ΦΔDF]?\s?\d[0-9X\-,\.]+\s?[EeΕεCc^M^m^Μ^μ^T^t^Τ^τ]?[ΚκKkMm^C^c]?\.?\s?"),
    }


    def __init__(self, title, debug = False):
        self.debug = debug
        self.info = self.find(title)                                                    


    def show_info(self):
        pprint(self.info)


    def get_data(self):
        list_ = []
        for col in self.cols:
            list_.append(self.info[col])
        return list_
        
    @staticmethod
    def get_columns():
        return ProductTitle.cols


    @staticmethod
    def remove_diacritics(title) -> str:
        d = {ord('\N{COMBINING ACUTE ACCENT}'):None}                                    # unicodedata library
        normalized_title = ud.normalize("NFD", title).upper().translate(d)              # code to remove diacritics
        return normalized_title


    @staticmethod
    def to_excel(data, column_names, start = False, filename="product_title_results.xlsx"):
        print("Creating excel file.")     
        df = pd.DataFrame(data, columns=column_names)                                           
        try:                                                                            #Creates an excel file with data
            df.to_excel(filename)                                                       #start=True will launch excel
            if start:
                print("Launching Excel File")
                os.startfile(filename)                                      
        except Exception as e:                                                                  
            print("An exception was raised:")
            print(e)
            print("Try closing excel and retry")
            inp = input("Retry? Y/n").lower()
            if inp != "n" or inp != "ν":
                df.to_excel(filename)
                if start:
                    print("Launching Excel File")
                    os.startfile(filename)


    #find() function is huge and needs to be simplified with more descriptive names and standarized output. Once everything
    #works as it should be, I will split it into multiple functions 

    def find(self, string: str) -> dict[str:str]:

                                                          
        og_title = string                                                           
        # print('\n'+string.strip())
        string = string + " |"
        string = string.upper()


        #remove info in parenthesis !!
        parenth = re.findall(self.regex_patterns['parenth'], string)                                   #the idea is that info in parenthesis
        if parenth:                                                                 #is just miscellaneous info
            for item in parenth:                                                    #and doesn't need to be looked upon too analytically
                string = string.replace(f" ({item})", "")
                string = re.sub("\s{2,}", " ", string)
            parenth = " ".join(parenth)
        else:
            parenth = ""

        #debrand !!
        brand = ""
        for brand_name in ProductTitle.topics['brand']:
            if fuzz.partial_ratio(string, brand_name.upper()) == 100:               #Brand names generally are very easy to isolate and could
                brand = brand_name                                                  #be handled with a simple search, but to avoid possible
                string = string.replace(brand.upper(), "").strip()                  #typos, I did a fuzz search instead
                break


        #deseries !!
        series = ""                                                                 #Some products are part of named series that usually
        words = string.split()[0:2]                                                 #come before the product. While it's not easy to isolate
        if re.search("\d+", words[1]) and re.search("[a-zA-Z]+", words[0]):         #them, if they follow a "Name ##" pattern, then they are 
            series = " ".join(words)                                                #almost certainly a match.
            string = string.replace(series + " ", "") 


        #degroup !!
        #patterns = ["([Ss])(/\d*)\s", "(\s\d\d?)?(\s[SΣsσ][ΕεEe][TtΤτ]\s\d{1,4}\s\S+)\s"]#, "([SΣsσ][ΕεEe][TtΤτ]).*?(\d\d?\s[TtΤτ][ΕεEe][MmΜμ]\.?)"]          #Poonto adds the tag 'ΣΕΤ # τεμ.' to certain products
        for pattern in self.regex_patterns['grouping']:                                                    #and unless there is a typo, this should always work fine.
            grouping = re.findall(pattern, string)
            if grouping:
                grouping = grouping[0] if type(grouping) != type("string") else grouping 
                for word in grouping:
                    if word and word != " ":
                        string = string.replace(word, "").strip()
                grouping = " ".join(list(grouping[0])).strip()
                break
            grouping = ""
        
        string = re.sub("\s{2,}", " ", string)

        
        ####### Alphanumeric split with pipes
        pattern = "(\S*\d\S*)\s?"
        string = re.sub(pattern, "| \\1 ", string).strip("|\n")


        #de-SKU
        code = ""
        check_string = string.split("|")[-1].strip()                                         
        if len(check_string) > 3:                                                       #SKUs have more than 3 characters
            found = re.search("\d[XxΧχ]\d", check_string)                               #if they match too closely to a dimension pattern
            if not found:                                                               #then the program should ignore them
                found = re.search("\S+\d\S+", check_string)
                if found:
                    if re.search("[a-zA-Z]", check_string):                             #if the SKU contains letters, it should never begin
                        if not re.search("^[a-zA-Z]", check_string):                 #with a numeric character.
                            found = ""                                                  
                if found:# and found.span()[1]-found.span()[0] > 2:
                    string = string.replace(check_string, "").strip().strip("|").strip()
                    string = string.replace("κωδ.", "").replace("ΚΩΔ.", "")
                    code = check_string


        volumes = ""
        pattern = "\| \d+,?\d?\s?[MmLlCc][TtLlCc\s]"
        results = re.findall(self.regex_patterns['volume'], string)
        string = re.sub(self.regex_patterns['volume'], "", string)
        for i, y in enumerate(results):
            results[i] = results[i].strip().strip("|").strip()
        volumes = " / ".join(results)
           
     
        # #possible dimensions
        # dimensions = ""
        # pattern = "\S+\s[EeΕε][ΚκKk]\.?\s|[ΦΔΥ]?\d\S*[ΧχXx]?\S*\s?[ΕΚεκCMcm]{0,2}\.?"
        # search_whole_string = re.findall(pattern, string)                                 #if the pattern repeats, then we can't know 
        # search = re.findall(pattern, string.split("|")[-1])                               #what product the dimension references
        # if len(search) == 1 and len(search_whole_string) == 1:                          #It's safe only if there is only a single match.
        #     dimensions = search[0]
        #     string = string.replace(dimensions, "").strip().strip("|").strip()
        #     dimensions = dimensions.lower().replace("χ", "x")


        dimensions: str = ""
        while re.search("(\d)[xΧχ](\d)", string):
            string = re.sub("(\d)[xΧχ](\d)", r"\1X\2", string)
        dimensions_pattern = "\| [ΦΔDF]?\s?\d[0-9X\-,\.]+\s?[EeΕεCc^M^m^Μ^μ^T^t^Τ^τ]?[ΚκKkMm^C^c]?\.?\s?"
        dimension_results = re.findall(self.regex_patterns['dimension'], string)
        if len(dimension_results) == 1:
            dimensions = dimension_results[0].strip().strip("|").strip()
            string = re.sub(f"\|\s{dimensions}", "", string)
            string = re.sub("\s{2,}", "", string)       


        rest = ""
                                                                                                        

        return {'og_title': og_title, 
                'Title':string.title().strip(),
                'Rest':rest.title().strip(),
                'Brand':brand.title().strip(), 
                'Grouping':grouping.title().strip(), 
                'SKU':code.upper().strip(),
                'Dimensions':dimensions, #.lower().strip(),
                'Series':series.title().strip(), 
                'Misc':parenth.title().strip(), 
                'Volume':volumes.title().strip(),
        }
 


#----------------------------------------------------------


if "dumped_product_titles.txt" in os.listdir():
    print("Loading Cache - To load from scratch, delete 'dumped_product_titles.txt'")               #loads cached data 
    with open("dumped_product_titles.txt", 'r', encoding='utf-8') as f:
        titles = f.readlines()
else:
    filename = "products_20221215-044713.xlsx"                                                      #loads original data
    print("Loading Workbook")                                                                       #for testing purposes, just use the cache
    wb = opxl.load_workbook(filename)                                                               #should change for final version
    sheet = wb["Products"]

    titles = [x[0].value for x in sheet["f2:f100000"] if x[0].value]
    with open("dumped_product_titles.txt", 'w', encoding='utf-8') as f:
        f.write("\n".join(titles))
        wb.close()


def test():
    pt = ProductTitle
    e = {}
    for i in range(50):
        no = randint(0,60000)
        # print(no)
        e[i] = pt(titles[no])
        print(f"{i} - {e[i].title}")
        print(e[i].brand, e[i].code, e[i].grouping, e[i].color, e[i].material, e[i].dimension, sep="\n")
        print(e[i].product)
        print()


def word_counter():
    longstring = " ".join(titles).replace("\xa0", "0")
    all_words = longstring.split()
    return Counter(all_words)



df = pd.DataFrame(word_counter().items()).sort_values(1, ascending=False)
def get_new_dataset():
    df.to_excel("dataset.xlsx", index=False)
    os.startfile("dataset.xlsx")
    inp = input("Close excel and press enter to continue")
    dfnew = pd.read_excel('dataset.xlsx')
    dfnew.to_csv("updated_dataset.csv")
    return dfnew






def test_find(likely = 0, modulus = 1, show_only = ""):
    alltitles = 0
    count = 0
    for i, title in enumerate(titles):
        if i%modulus == 0:
            alltitles += 1
            product = ProductTitle(title).info        
            if likely:
                if re.search("\|", product['Title']) or re.search("\|", product["Rest"]): continue
            print("\n---\n")
            count += 1
            for k, v in product.items():
                if v:
                    if not show_only:
                        print(f"{k}: {v}")
                    else:
                        if product[show_only]:
                            print(f"{k}: {v}") 
    if alltitles:
        print(f"\n{alltitles = } / {count = } / {count/alltitles}")




def go(app, test=0, likely=0, modulus=1000, show_only=""):                          #Helper function for interactive shell            
    from importlib import reload                                                    #Reloads the module and runs a test function
    reload(app)                                                                     #Simple usage: app.go(app, test=0)
    if test:
        test_find(likely=likely, modulus=modulus, show_only=show_only)

pt = ProductTitle

def main():
    titles_count = len(titles)
    data = []
    print("Analyzing titles:")
    count = -1
    for i, title in enumerate(titles):
        # if i % int(len(titles)/10) == 0:
        #     count += 1
        # print(f"[{'|'*count}{' '*(10-count)}]", end='\r')
        data.append(ProductTitle(title).get_data())
    print(f"[{titles_count} / {titles_count}]")
    print("Analysis complete.")
    ProductTitle.to_excel(data, ProductTitle.get_columns(), start=True)






if __name__ == "__main__":
    start = perf_counter()
    main()
    print(perf_counter()-start)
