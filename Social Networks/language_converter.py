
class CyrilicToLatin:
    CYRILIC_ALPHABET = ['а', 'б', 'в', 'г', 'д', 'ђ', 'е', 'ж', 'з', 'и', 'ј', 'к', 
                    'л', 'љ', 'м', 'н', 'њ', 'о', 'п', 'р', 'с', 'т', 'ћ', 'у', 
                    'ф', 'х', 'ц', 'ч', 'џ', 'ш', 'А', 'Б', 'В', 'Г', 'Д', 'Ђ',
                    'Е', 'Ж', 'З', 'И', 'Ј', 'К', 'Л', 'Љ', 'М', 'Н', 'Њ', 'О',
                    'П', 'Р', 'С', 'Т', 'Ћ', 'У', 'Ф', 'Х', 'Ц', 'Ч', 'Џ', 'Ш']
    LATIN_ALPHABET = ['a', 'b', 'v', 'g', 'd', 'đ', 'e', 'ž', 'z', 'i', 'j', 'k',
                    'l', 'lj', 'm', 'n', 'nj', 'o', 'p', 'r', 's', 't', 'ć', 'u', 
                    'f', 'h', 'c', 'č', 'dž', 'š', 'A', 'B', 'V', 'G', 'D', 'Đ', 
                    'E', 'Ž', 'Z', 'I', 'J', 'K', 'L', 'Lj', 'M', 'N', 'Nj', 'O',
                    'P', 'R', 'S', 'T', 'Ć', 'U', 'F', 'H', 'C', 'Č', 'Dž', 'Š']

    CYRILIC_LATIN_ALPHABET = {}
    for letter in CYRILIC_ALPHABET:
        CYRILIC_LATIN_ALPHABET[letter] = LATIN_ALPHABET[CYRILIC_ALPHABET.index(letter)]
    
    def convertCyrilicToLatin(string):
        convStr = ""
        if (string is None):
            return convStr
        if (type(string)) is not str:
            return string
        for letter in string:
            latinLetter = CyrilicToLatin.CYRILIC_LATIN_ALPHABET.get(letter, letter)
            convStr += latinLetter
        return convStr