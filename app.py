import openpyxl as opxl
import pandas as pd
import sys, os
import unicodedata as ud
import re
from random import randint
from pprint import pprint as pp
from fuzzywuzzy import fuzz

class ProductTitle:

    topics: dict[str: re.Pattern] = {
    "color": ['-?ΛΕΥΚ\w-?', '-?ΑΣΠΡ\w-?', '-?ΓΚΡΙ-?', 
            '-?ΕΚΡΟΥ-?', '-?ΜΠΕΖ-?', '-?ΜΑΥΡ\w-?', 
            "-?ΑΝΘΡΑΚΙ-?", '-?ΚΟΚΚΙΝ\w-?', "-?ΜΠΟΡΝΤΩ-?", 
            "-?ΡΟΖ-?", "-?ΜΩΒ-?", '-?ΚΙΤΡΙΝ\w\w?-?', 
            '-?ΠΡΑΣΙΝ\w\w?-?', "-?ΦΥΣΤΙΚΙ-?", "-?ΒΕΡΑΜΑΝ-?", 
            '-?ΜΠΛΕ-?', '-?ΣΙΕΛ-?', 'ΓΑΛΑΖΙ\w', 
            '-?ΑΣΗΜΙ-?', '-?ΧΡΥΣ\w\w?-?', '-?ΜΟΥΣΤΑΡΔΙ-?',
            "-?ΠΕΤΡΟΛ-?","Σ?Κ?Ο?Υ?Ρ?Ο?\s?-?ΚΑΦΕ-?", "-?ΛΑΔΙ-?",
            "-?ΣΑΜΠΑΝΙ-?"],
    "brand": ["INART", "ESPIEL", "KENTIA", "ZAROS", "AI DECORATION", "CLICK", "GUY_LAROCHE", "SAINTCLAIR", "SB_HOME", "SBABY", "BLE"],
    "grouping": ["ΣΕΤ \d\d?\s?\S*", "ΣΕΤ \d\d?\s?\S*",  "ΣΕΤ ΤΩΝ \d\d?", "ΣΕΤ\d\d?", "SET", "^ΣΕΤ\s", "ΣΕΤ\s", "\d\d?\sΤΕΜ\S*", '\s(TEM)\s', "\s(ΤΕΜ)\s"],
    "dimension": ["[ΦΔDF]?\d{1,2},?\d?X\d{1,2},?\d?", "[ΦΔDF]?\d{1,2},?\d?Χ\d{1,2},?\d?", "[^0+](\d{1,2},?\d?)\s"],
    "unit": ["CM\w*\W?", "ΕΚ\w*\W?", "ΜΕΤΡ\w?\W?", "ML"],
    "sku": ["ΚΩΔ:\S+\d", "\S+\d"],
    "material": ['ΞΥΛ?Ι?Ν?\w?\w?', 'ΜΕΤΑΛΛ?Ι?Κ?\w?\w?', "ΨΑΘΙΝ\w\w?", "POLYRESIN", "ΠΟΛΥΕΣΤΕΡ\w?", "POLYESTER", "ΠΟΛΥΡΕΖΙΝ", "ΡΗΤΙΝΗ\w", "ΓΥΑΛ[^Α]\w*", "ΚΕΡΑΜΙΚ\w\w?", "ΥΦΑΣΜ?Α?Τ?Ι?Ν?\w\w?", "ΒΕΛΟΥΔΙ?Ν?\w?\w?", "FIBERGLASS"],
    }

    def __init__(self, title, debug = False):
        self.debug = debug
        self.original_title: str = title
        self.normalized_title: str = self.normalize()
        self.normalized_words = self.normalized_title.split(" ")
        self.matched_words = self.classifier()
        self.verbose_title = self.add_descriptors()
        self.product = self.extract_product()
        self.brand = self.extract("brand")
        self.code = self.extract("code")
        self.unit = self.extract("unit")
        self.grouping = self.extract("grouping")
        self.dimension = self.extract("dimension")
        self.color = self.extract("color")
        self.material = self.extract("material")
        self.entropy_title = self.calculate_entropy(self.normalized_title)
        self.entropy_product = self.calculate_entropy(self.product)


    def get_info(self) -> list:
        return [*map(lambda x: " ".join(x) if repr(type(x)) == "<class 'list'>" else x, [self.original_title, self.normalized_title, self.verbose_title, self.product, self.brand, self.code, self.unit, self.grouping, self.dimension, self.color, self.material, self.entropy_product, self.entropy_title])]

    
    def get_column_names(self) -> list[str]:
        return ["og_title", "normalized_title", "verbose_title", "product", "brand", "code", "unit", "grouping", "dimension", "color", "material", "entropy", "entropy_title"]


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
                        if check:
                            print(check)
                    if re.search(pattern, word):
                        matched_words[word] = topic
                        is_matched = True
                        break
                if is_matched: break    
        return matched_words


    def classifier(self) -> bool:
        matched = {}
        for topic, patterns in self.topics.items():
            for pattern in patterns:
                search = re.findall(pattern, self.normalized_title)
                if search:
                    for item in search:
                        matched[item] = topic
        if self.debug:
            print(self.original_title)
            print(matched)
        return matched



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
            .replace("+ ", "+")
            .replace("ΚΩΔ ", "ΚΩΔ:")
            .replace("ΣΕΤ ΤΩΝ", "ΣΕΤ")
            .replace("AI DECORATION", "AI_DECORATION")
            .replace("SB HOME", "SB_HOME")
            .replace("GUY LAROCHE", "GUY_LAROCHE")
        )
        normalized_title = re.sub("\s{2,}", " ", normalized_title)
        return normalized_title

    def add_descriptors(self):
        title = self.normalized_title
        for word, topic in self.matched_words.items():
            title = title.replace(word, f"${topic[0]}{word}")
        return title


    def extract_product(self) -> str:
        product = self.verbose_title
        product = re.sub("\$\$\$\S+\)\s?", "", product)
        return product.replace("  ", " ").replace("  ", " ")


    def extract(self, kw) -> list[str]:
        return re.findall(f"\$\$\$(\S+)\({kw}\)", self.verbose_title)


    def calculate_entropy(self, text: str) -> float:
        words = re.findall("\S+", text)
        score = 0
        for word in words:
            score += 0.01
            if len(word) <= 3:
                score += 1/len(word)
        return score
        

        



def main():

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

    data = []
    for i, title in enumerate(titles):
        title = ProductTitle(title.strip())
        data.append(title.get_info())
        print(f"[{i+1}/{len(titles)}]                ", end="\r")

    print("Preparing Excel File")
    title.to_excel(data, title.get_column_names(), start=True)





if __name__ == "__main__":
    main()
