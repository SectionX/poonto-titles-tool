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

class ProductTitle:

    topics: dict[str: re.Pattern] = {
    "color": ['[^Μ][^Ε]\s?-?(ΛΕΥΚ\w[Σ]?)-?', 
              '[^Μ][^Ε]\s?-?(ΑΣΠΡ\w[Σ]?)-?', 
              '[^Μ][^Ε]\s?-?(ΓΚΡΙ)-?', 
              '[^Μ][^Ε]\s?-?(ΕΚΡΟΥ)-?', 
              '[^Μ][^Ε]\s?-?(ΜΠΕΖ)-?', 
              '[^Μ][^Ε]\s?-?(ΜΑΥΡ\w[Σ]?)-?', 
              "[^Μ][^Ε]\s?-?(ΑΝΘΡΑΚΙ)-?", 
              '[^Μ][^Ε]\s?-?(ΚΟΚΚΙΝ\w[Σ]?)-?', 
              "[^Μ][^Ε]\s?-?(ΜΠΟΡΝΤΩ)-?", 
              "[^Μ][^Ε]\s?-?(ΡΟΖ)-?", 
              "[^Μ][^Ε]\s?-?(ΜΩΒ)-?", 
              '[^Μ][^Ε]\s?-?(ΚΙΤΡΙΝ\w[Σ]?)-?', 
              '[^Μ][^Ε]\s?-?(ΠΡΑΣΙΝ\w[Σ]?)-?', 
              "[^Μ][^Ε]\s?-?(ΦΥΣΤΙΚΙ)-?", 
              "[^Μ][^Ε]\s?-?(ΒΕΡΑΜΑΝ)-?", 
              '[^Μ][^Ε]\s?-?(ΜΠΛΕ)-?', 
              '[^Μ][^Ε]\s?-?(ΣΙΕΛ)-?', 
              '[^Μ][^Ε]\s?-?(ΓΑΛΑΖΙ\w[Σ]?)-?', 
              '[^Μ][^Ε]\s?-?(ΑΣΗΜΙ)-?', 
              '[^Μ][^Ε]\s?-?(ΧΡΥΣ\w[Σ]?)-?', 
              '[^Μ][^Ε]\s?-?(ΜΟΥΣΤΑΡΔΙ)-?',
              "[^Μ][^Ε]\s?-?(ΠΕΤΡΟΛ)-?",
              "[^Μ][^Ε]\s?-?(ΣΚΟΥΡΟ)-?\s?", 
              "[^Μ][^Ε]\s?-?(ΚΑΦΕ)-?", 
              "[^Μ][^Ε]\s?-?([^K]?ΛΑΔΙ)-?",
              "[^Μ][^Ε]\s?-?(ΣΑΜΠΑΝΙ)-?", 
              "[^Μ][^Ε]\s?-?(NATURAL)-?", 
              "[^Μ][^Ε]\s?-?(ΔΙΑΦΑΝ\w[Σ]?)-?", 
              "[^Μ][^Ε]\s?-?(ΚΡΕΜ)\W", 
              "[^Μ][^Ε]\s?-?(Τ\wΡΚΟΥΑΖ)-?", 
              "[^Μ][^Ε]\s?-?(ΜΠΡΟΝΖΕ)-?"],
    "brand": ["INART", "ESPIEL", "KENTIA", "ZAROS", "AI DECORATION", "CLICK", "GUY LAROCHE", "SAINT CLAIR", "SAINTCLAIR", "SB HOME", "SBABY", "BLE", "Versace 19•69"],
    "grouping": ["ΣΕΤ \d\d?\s?\S*", "ΣΕΤ \d\d?\s?\S*",  "ΣΕΤ ΤΩΝ \d\d?", "ΣΕΤ\d\d?", "SET", "^ΣΕΤ\s", "ΣΕΤ\s", "\d\d?\sΤΕΜ\S*", '\s(TEM)\s', "\s(ΤΕΜ)\s", "\sS\s\d\d?", "^S\s\d\d?"],
    "dimension": ["[ΦΔDF]\d\S+\s?\d?\d?\d?\s?[CM,ML,L,ΕΚ]*", "\S*\d[XΧ]\d\S*\s?\d?\d?\d?\s?[CM,ML,L,ΕΚ]*", "\S*\d\s?\d?\d?\d?\s?[CMLΕΚΧΙΛ]+\s"],
    # "unit": ["CM\w*\W?", "ΕΚ\w*\W?", "ΜΕΤΡ\w?\W?", "ML"],
    # "sku": ["ΚΩΔ:\S+\d", "\S+\d"],
    "material": ["ΑΛΟΥΜΙΝ\S*", "COTTON", "ΠΟΡΣΕΛΑΝ\S*", "ΜΠΑΜΠΟΥ", "BAMBOO", 'ΞΥΛ\S*', 'ΜΕΤΑΛΛ?Ι?Κ?\w?\w?', "ΨΑΘΙΝ\w\w?", "POLYRESIN", "ΠΟΛΥΕΣΤΕΡ\w?", "POLYESTER", "ΠΟΛΥΡΕΖ\S*", "ΡΗΤΙΝΗ\w", "ΓΥΑΛ[^Α]\S*", "ΚΕΡΑΜΙΚ\w\w?", "ΥΦΑΣΜ?Α?Τ?Ι?Ν?\w\w?", "ΒΕΛΟΥΔΙ?Ν?\w?\w?", "FIBERGLASS"],
    }

    def __init__(self, title, debug = False):
        self.debug = debug
        self.original_title: str = title.strip()
        self.title = self.original_title
        self.normalized_title, self.brand, self.code = self.simplify_title()
        self.grouping = self.classifier("grouping")
        self.color = self.classifier("color")
        self.material = self.classifier("material")
        self.dimension = self.classifier("dimension")
        self.product = self.extract_product().title()
        self.entropy_title = self.calculate_entropy(self.normalized_title)
        self.entropy_product = self.calculate_entropy(self.product)





        # self.normalized_words = self.normalized_title.split(" ")
        # self.matched_words = self.classifier()
        # self.verbose_title = self.add_descriptors()
        # self.product = self.extract_product()
        # # self.brand = self.extract("brand")
        # # self.code = self.extract("code")
        # self.unit = self.extract("unit")
        # self.grouping = self.extract("grouping")
        # self.dimension = self.extract("dimension")
        # self.color = self.extract("color")
        # self.material = self.extract("material")


    def simplify_title(self):
        title = self.normalize()
        brand_ = ""
        possible_SKU = ""
        for brand in self.topics['brand']:
            if fuzz.partial_ratio(title, brand) == 100:
                brand_ = brand
                title = title.replace(brand, "").strip()
                break
        possible_SKU = title.split(" ")[-1]
        if re.search("\d", possible_SKU):
            title = title.replace(possible_SKU, "")
        else:
            possible_SKU = ""
        if self.debug:
            print(possible_SKU)
            print(brand_)
            print(title)
            print()
        # self.brand = brand_
        # self.code = possible_SKU
        return title, brand_, possible_SKU.replace("ΚΩΔ:", "")


    def normalize(self) -> str:
        d = {ord('\N{COMBINING ACUTE ACCENT}'):None}                                            # unicodedata library
        normalized_title = ud.normalize("NFD", self.original_title).upper().translate(d)        # code to remove diacritics
        normalized_title = (
            normalized_title
            .replace(" ", " ")
            .replace(" ", " ")
            .replace("/", " ")
            # .replace("-", " ")
            .replace("\"", "")
            .replace("\'", "")
            .replace(", ", " ")
            .replace("(", " ")
            .replace(")", " ")
            .replace(".", "")
            .replace(",1", ".1")
            .replace(",2", ".2")
            .replace(",3", ".3")
            .replace(",4", ".4")
            .replace(",5", ".5")
            .replace(",6", ".6")
            .replace(",7", ".7")
            .replace(",8", ".8")
            .replace(",9", ".9")
            .replace(",", " ")
            .replace(".1", ",1")
            .replace(".2", ",2")
            .replace(".3", ",3")
            .replace(".4", ",4")
            .replace(".5", ",5")
            .replace(".6", ",6")
            .replace(".7", ",7")
            .replace(".8", ",8")
            .replace(".9", ",9")
            # .replace(" ΤΕΜ", "ΤΕΜ")
            # .replace("+ ", "+")
            .replace("ΚΩΔ ", "ΚΩΔ:")
            .replace("ΣΕΤ ΤΩΝ", "ΣΕΤ")
            .replace("ΧΡΩΜΑΤΑ", "")
            .replace("ΧΡΩΜΑ", "")
            .replace(" ΣΕ ", " ")
            # .replace("AI DECORATION", "AI_DECORATION")
            # .replace("SB HOME", "SB_HOME")
            # .replace("GUY LAROCHE", "GUY_LAROCHE")
        )
        normalized_title = re.sub("\s{2,}", " ", normalized_title)
        normalized_title = re.sub("\d*%", "", normalized_title)
        return normalized_title


    def get_info(self) -> list:
        return [*map(lambda x: " ".join(x).title() if repr(type(x)) == "<class 'list'>" else str(x).title(), [self.original_title, self.normalized_title, self.product, self.brand, self.code, self.grouping, self.dimension, self.color, self.material])]+[self.entropy_product, self.entropy_title]

    
    def get_column_names(self) -> list[str]:
        return ["og_title", "normalized_title", "product", "brand", "code", "grouping", "dimension", "color", "material", "entropy", "entropy_title"]


    def to_excel(self, data, column_names, start = False):
        df = pd.DataFrame(data, columns=column_names)
        try:
            df.to_excel("product_title_results.xlsx")
            if start:
                print("Launching Excel File")
                os.startfile("product_title_results.xlsx")
        except Exception as e:
            print("An exception was raised:")
            print(e)
            print("Try closing excel and retry")
            inp = input("Retry? Y/n").lower()
            if inp != "n" or inp != "ν":
                df.to_excel("product_title_results.xlsx")
                if start:
                    print("Launching Excel File")
                    os.startfile("product_title_results.xlsx")



    def classifier_old(self) -> bool:
        matched_words = {}
        for word in self.normalized_words:
            is_matched = False
            matched_words.setdefault(word, "")
            for topic, patterns in self.topics.items():
                for pattern in patterns:
                    if self.debug:
                        check = re.search(pattern, word)
                        # if check:
                        #     print(check)
                    if re.search(pattern, word):
                        matched_words[word] = topic
                        is_matched = True
                        break
                if is_matched: break    
        return matched_words


    def classifier(self, topic) -> bool:
        matched = {}
        patterns = self.topics[topic]
        for pattern in patterns:
            search = re.findall(pattern, self.normalized_title)
            if search:
                for item in search:
                    item = item.strip()
                    matched[item] = topic
        if self.debug:
            print(self.original_title)
            print(matched)
        return list(set(list(matched.keys())))


    def add_descriptors(self):
        title = self.normalized_title
        for word, topic in self.matched_words.items():
            title = title.replace(word, f"${topic[0]}{word}")
        return title


    def extract_product(self) -> str:
        product = self.normalized_title+" "
        items = [self.dimension, self.grouping, self.material, self.color]
        string = ""
        for item in items:
            word = " ".join(item)
            string += f" {word}"
        string = re.sub("\s{1,}", " ", string)
        words = string.split(" ")
        if self.debug:
            print(words)
        for word in words:
            if word:
                # word = word.replace("&&&", " ")
                product = product.replace(f" {word} ", "  ")
        return re.sub("\s{1,}", " ", product).strip()


    def extract(self, kw) -> list[str]:
        return re.findall(f"\$\$\$(\S+)\({kw}\)", self.verbose_title)


    def calculate_entropy(self, text: str) -> float:
        words = re.findall("\S+", text)
        score = 0
        for word in words:
            score += 0.01
            if len(word) <= 3:
                score += 1.01
        return score
        


if "dumped_product_titles.txt" in os.listdir():
    print("Loading Cache - To load from scratch, delete 'dumped_product_titles.txt'")
    with open("dumped_product_titles.txt", 'r', encoding='utf-8') as f:
        titles = f.readlines()
else:
    filename = "products_20221215-044713.xlsx"
    print("Loading Workbook")
    wb = opxl.load_workbook(filename)
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



def find(string: str = titles[1231], brand = "", grouping = "", code = "", first: int = 1) -> tuple[str, str, str, int]:


    if first:
        string = string + " |"

        #debrand
        for brand_name in ProductTitle.topics['brand']:
            if fuzz.partial_ratio(string, brand_name) == 100:
                brand = brand_name
                string = string.replace(brand, "").strip()
                break

        #degroup
        pattern = "([ΣΕΤσετSETset]{3}\s\d{1,4}\s\S+)\s"
        grouping = re.findall(pattern, string)
        if grouping:
            grouping = grouping[0]
            string = string.replace(grouping+" ", "").strip()
        
        first = 0

    head, *tail = string.split("|")

    # recursive call to split string into parts with alphanumerics
    pattern = "(.*)\s(\w*\d.*)\s?"
    results = re.findall(pattern, head)
    if results:
        results = list(results[0])
        head = " | ".join(results)
        return find(head + "|" + "|".join(tail), brand, grouping, code, 0)
    string = string.strip("|")

    #de-SKU
    check_string = string.split("|")[-1]
    if len(check_string) > 3:
        found = re.search("\d[XxΧχ]\d", check_string)
        if not found:
            found = re.search("\S+\d\S+", check_string)
            if found:# and found.span()[1]-found.span()[0] > 2:
                string = string.replace(check_string, "")
                string = string.replace("κωδ.", "")
                code = check_string.strip().strip("|").strip()


    return(string.title().strip().strip("|").strip(), brand.title(), grouping, code)


    


def main():


    data = []
    for i, title in enumerate(titles):
        title = ProductTitle(title.strip())
        data.append(title.get_info())
        print(f"[{i+1}/{len(titles)}]                ", end="\r")

    print("Preparing Excel File")
    title.to_excel(data, title.get_column_names(), start=True)





if __name__ == "__main__":
    main()
