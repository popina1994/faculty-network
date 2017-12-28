
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

    def convertCyrilicToLatin(str):
        convStr = ""
        for letter in str:
            index = -1
            try:
                index = CyrilicToLatin.CYRILIC_ALPHABET.index(letter)
            except:
                index = -1
            if (index != -1):
                latinLetter = CyrilicToLatin.LATIN_ALPHABET[CyrilicToLatin.CYRILIC_ALPHABET.index(letter)]
            else:
                latinLetter = letter
            convStr += latinLetter
        return convStr