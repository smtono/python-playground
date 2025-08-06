def reverse_word(word):
    reversed = ""
    for letter in word:
        reversed = letter + reversed
    return reversed

if __name__ == "__main__":
    word = input("Word: ")
    reverse_word(word)