def reverse_word(word):
    rev = ""
    for letter in word:
        rev = letter + rev
    return rev


if __name__ == "__main__":
    word = input("Word: ")
    reverse_word(word)
