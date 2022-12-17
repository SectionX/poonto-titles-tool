import openpyxl as opxl
import pandas as pd
import sys, os
import unicodedata as ud
import re
from random import randint
from pprint import pprint as pp

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
    "dimension": ["[ΦΔDF]?\d{1,2},?\d?X\d{1,2},?\d?", "[ΦΔDF]?\d{1,2},?\d?Χ\d{1,2},?\d?", "[^0+]\d{1,2},?\d?"],
    "grouping": ["ΣΕΤ\d\d?", "SET", "^ΣΕΤ", "ΣΕΤ", "\d*ΤΕΜ\w*\.?"],
    "unit": ["CM\w*\W?", "ΕΚ\w*\W?", "ΜΕΤΡ\w?\W?", "ML"],
    "code": ["ΚΩΔ:\S+\d", "\S+\d"],
    "material": ['ΞΥΛ?Ι?Ν?\w?\w?', 'ΜΕΤΑΛΛ?Ι?Κ?\w?\w?', "ΨΑΘΙΝ\w\w?", "POLYRESIN", "ΠΟΛΥΕΣΤΕΡ\w?", "POLYESTER", "ΠΟΛΥΡΕΖΙΝ", "ΡΗΤΙΝΗ\w", "ΓΥΑΛ[^Α]\w*", "ΚΕΡΑΜΙΚ\w\w?", "ΥΦΑΣΜ?Α?Τ?Ι?Ν?\w\w?", "ΒΕΛΟΥΔΙ?Ν?\w?\w?", "FIBERGLASS"],
    }

    def __init__(self, title):
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


    def get_info(self) -> list:
        return [self.original_title, self.verbose_title, self.product, self.brand, self.code, self.unit, self.grouping, self.dimension, self.color, self.material]

    
    def get_column_names(self) -> list[str]:
        return ["og_title", "verbose_title", "product", "brand", "code", "unit", "grouping", "dimension", "color", "material"]


    def to_excel(self, data, column_names, start = False):
        df = pd.DataFrame(data, columns=column_names)
        df.to_excel("product_title_results.xlsx")
        if start:
            os.startfile("product_title_results.xlsx")


    def classifier(self) -> bool:
        matched_words = {}
        for word in self.normalized_words:
            is_matched = False
            matched_words.setdefault(word, "")
            for topic, patterns in self.topics.items():
                for pattern in patterns:
                    if re.match(pattern, word):
                        matched_words[word] = topic
                        is_matched = True
                        break
                if is_matched: break    
        return matched_words


    def normalize(self) -> str:
        d = {ord('\N{COMBINING ACUTE ACCENT}'):None}                                            # unicodedata library
        normalized_title = ud.normalize("NFD", self.original_title).upper().translate(d)        # code to remove diacritics
        normalized_title = (
            normalized_title
            .replace(" ", " ")
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
            .replace(" ΤΕΜ", "ΤΕΜ")
            .replace("+ ", "+")
            .replace("ΚΩΔ ", "ΚΩΔ:")
            .replace("ΣΕΤ ΤΩΝ", "ΣΕΤ")
            .replace("AI DECORATION", "AI_DECORATION")
            .replace("SB HOME", "SB_HOME")
            .replace("GUY LAROCHE", "GUY_LAROCHE")
        )
        return normalized_title

    def add_descriptors(self):
        title = self.normalized_title
        for word in self.normalized_words:
            if self.matched_words[word]:
                title = title.replace(word, f"$$${word}({self.matched_words[word]})")
        return title


    def extract_product(self) -> str:
        product = self.verbose_title
        product = re.sub("\$\$\$\S+\)\s?", "", product)
        return product.replace("  ", " ").replace("  ", " ")


    def extract(self, kw) -> list[str]:
        return re.findall(f"\$\$\$(\S+)\({kw}\)", self.verbose_title)




def main():

    filename = "products_20221215-044713.xlsx"
    wb = opxl.load_workbook(filename)
    sheet = wb["Products"]

    titles = [x[0].value for x in sheet["f2:f100000"] if x[0].value]
    wb.close()

    data = []
    for title in titles:
        title = ProductTitle(title)
        data.append(title.get_info())


    title.to_excel(data, title.get_column_names(), start=True)



if __name__ == "__main__":
    main()
