
import win32com.client as wincl
import random
import sys
    
speak = wincl.Dispatch("SAPI.SpVoice")

def speak_and_print(text):
    print(text)
    speak.Speak(text)

wordlist = []

with open(r"C:/Users/Dave/Documents/yr34_spellings.txt") as words:
    for word in words:
        wordlist.extend(word.split())

welcome = "Welcome to super speller! Would you like to play?"
how_many_questions_text = "would you like to answer 10, 15 or 20 questions?"
how_do_you_spell = 'How do you spell... '

speak_and_print(welcome)
print('Enter y or n ')

if input()[0] != 'y':
    speak_and_print('Goodbye!')
    sys.exit()
else:
    speak_and_print("Great! Let's play")
    
while True:
    try:
        speak_and_print(how_many_questions_text)
        number_of_questions = int(input())
        if number_of_questions not in [10,15,20]:
            raise ValueError
    except ValueError:
        print('That is not a valid option. Try again or press q to quit')
    else:
        break

speak_and_print('You selected ' + str(number_of_questions) + ' questions. Starting in...')

score = 0

for i in range(5,0,-1):
    speak_and_print(str(i))

chosen_words = random.sample(wordlist,number_of_questions)
chosen_words_guesses = []

for q in range(number_of_questions):
    speak_and_print('Question ' + str(q+1) + ') ' + how_do_you_spell)
    speak.Speak(chosen_words[q])
    guess = input()
    chosen_words_guesses.append(guess)
    
print("")
print("RESULTS")
print("")

for i,word in enumerate(chosen_words):
    if word == chosen_words_guesses[i]:
        score = score + 1
        
print("")

superlatve = 'Better luck next time! '
if (score / number_of_questions > 0.5):
    superlative = 'Not bad. '
if (score / number_of_questions > 0.75):
    superlative = 'Well done. '
if (score / number_of_questions > 0.9):
    superlative = 'AMAZING. '    

speak_and_print(superlative + 'You scored ' + str(score) + ' out of ' + str(number_of_questions))

if (score != number_of_questions):
    for q,word in enumerate(chosen_words):
        if word != chosen_words_guesses[q]:
            speak_and_print('Question ' + str(q+1) + ') ' + 'The word was ' + word + ' (' + ' '.join(list(word.upper())) + '). You wrote ' + chosen_words_guesses[q])
