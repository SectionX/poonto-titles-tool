class Product:

    cols = [
        'og_title',
        'Title',
        'Rest',
        'Brand',
        'Grouping',
        'SKU',
        'Dimensions',
        'Series',
        'Misc',
        'Volume',
    ]

    topics: dict[str: str] = {
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
    "brand": ["INART", "ESPIEL", "KENTIA", "ESTIA", "ΕΣΤΙΑ", "ZAROS", "AI DECORATION", "CLICK", "GUY LAROCHE", "SAINT CLAIR", "SAINTCLAIR", "SB HOME", "SBABY", "BLE", "Versace 19•69"],
    "grouping": ["ΣΕΤ \d\d?\s?\S*", "ΣΕΤ \d\d?\s?\S*",  "ΣΕΤ ΤΩΝ \d\d?", "ΣΕΤ\d\d?", "SET", "^ΣΕΤ\s", "ΣΕΤ\s", "\d\d?\sΤΕΜ\S*", '\s(TEM)\s', "\s(ΤΕΜ)\s", "\sS\s\d\d?", "^S\s\d\d?"],
    "dimension": ["[ΦΔDF]\d\S+\s?\d?\d?\d?\s?[CM,ML,L,ΕΚ]*", "\S*\d[XΧ]\d\S*\s?\d?\d?\d?\s?[CM,ML,L,ΕΚ]*", "\S*\d\s?\d?\d?\d?\s?[CMLΕΚΧΙΛ]+\s"],
    # "unit": ["CM\w*\W?", "ΕΚ\w*\W?", "ΜΕΤΡ\w?\W?", "ML"],
    # "sku": ["ΚΩΔ:\S+\d", "\S+\d"],
    "material": ["ΑΛΟΥΜΙΝ\S*", "COTTON", "ΠΟΡΣΕΛΑΝ\S*", "ΜΠΑΜΠΟΥ", "BAMBOO", 'ΞΥΛ\S*', 'ΜΕΤΑΛΛ?Ι?Κ?\w?\w?', "ΨΑΘΙΝ\w\w?", "POLYRESIN", "ΠΟΛΥΕΣΤΕΡ\w?", "POLYESTER", "ΠΟΛΥΡΕΖ\S*", "ΡΗΤΙΝΗ\w", "ΓΥΑΛ[^Α]\S*", "ΚΕΡΑΜΙΚ\w\w?", "ΥΦΑΣΜ?Α?Τ?Ι?Ν?\w\w?", "ΒΕΛΟΥΔΙ?Ν?\w?\w?", "FIBERGLASS"],
    }