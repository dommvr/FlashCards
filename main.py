import pandas as pd
import pyttsx3
import random
import speech_recognition as sr
import os


clear = lambda: os.system('cls')


class Flashcards:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)
        self.r = sr.Recognizer()
        self.mic = sr.Microphone()
        self.polish = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_PL-PL_PAULINA_11.0'
        self.english = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0'
        self.voices = {
            "English": ["HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0", "en-US"],
            "Polish": ["HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_PL-PL_PAULINA_11.0", "pl-PL"]
        }
        self.df_words = pd.read_excel('words.xlsx')

    def say_text(self, language, text):
        self.engine.setProperty('voice', language)
        self.engine.say(text)
        self.engine.runAndWait()

    def menu(self):
        clear()
        print("Select learning method: \nLearn by listening and reading words [1]\nLearn by saying words [2]\nLearn by writing words [3]")
        while True:
            try:
                self.met_choice = int(input('Choice: '))
            except ValueError:
                continue
            if self.met_choice == 1:
                print('')
                FC.set_language()
                FC.select_category()
                FC.main_loop_listening()
            elif self.met_choice == 2:
                FC.select_language()
                FC.select_category()
                FC.main_loop_saying()
            elif self.met_choice == 3:
                FC.select_language()
                FC.select_category()
                FC.main_loop_writing()

    def set_language(self):
        self.sel_language = self.df_words.columns[0]
        self.sel_voice = self.voices[self.sel_language][0]
        self.sec_language = self.df_words.columns[1]
        self.sec_voice = self.voices[self.sec_language][0]

    def select_language(self):
        print(f"\nSelect language to learn\n{self.df_words.columns[0]} [1]\n{self.df_words.columns[1]} [2]")
        while True:
            try:
                self.lan_choice = int(input('Choice: '))
            except ValueError:
                continue
            if self.lan_choice == 1:
                self.sel_language = self.df_words.columns[1]
                self.sec_language = self.df_words.columns[0]
                break
            elif self.lan_choice == 2:
                self.sel_language = self.df_words.columns[0]
                self.sec_language = self.df_words.columns[1]
                break
        self.sel_voice = self.voices[self.sel_language][0]
        self.sec_voice = self.voices[self.sec_language][0]
        self.rec_language = self.voices[self.sec_language][1]

    def select_category(self):
        self.categories = list(set(self.df_words['Category']))
        self.counter = 1
        print('\nSelect category to learn')
        for category in self.categories:
            print(f"{category} [{self.counter}]")
            self.counter += 1
        print(f"All categories [{self.counter}]")
        while True:
            try:
                self.category = int(input('Choice: '))
                if 1 <= self.category <= len(self.categories)+1:
                    break
            except ValueError:
                continue
        if self.category == len(self.categories)+1:
            self.sel_cat_words = self.df_words.iloc[:, 0:2]
        else:
            self.sel_category = self.categories[self.category-1]
            self.sel_cat_words = self.df_words.loc[self.df_words['Category'] == self.sel_category]
            self.sel_cat_words = self.sel_cat_words.reset_index(drop=True)

    def main_loop_listening(self):
        self.words_index = list(range(len(self.sel_cat_words.index)))
        random.shuffle(self.words_index)
        print('\nCommands:\n-next (skip flashcard)\n-repeat (repeat flashcard)\n-menu (go back to menu)\n-exit (exit program)\n')
        for index in self.words_index:
            print(f"{self.sel_language}: {self.sel_cat_words.at[index, self.sel_language]}\n{self.sec_language}: {self.sel_cat_words.at[index, self.sec_language]}")
            while True:
                FC.say_text(self.sel_voice, self.sel_cat_words.at[index, self.sel_language])
                FC.say_text(self.sec_voice, self.sel_cat_words.at[index, self.sec_language])
                FC.get_command(['-next', '-repeat', '-menu', '-exit'])
                if self.command == '-next':
                    break
                elif self.command == '-repeat':
                    continue
                elif self.command == '-menu':
                    FC.menu()
                elif self.command == '-exit':
                    exit()
        FC.end_menu()

    def get_command(self, commands):
        while True:
            self.command = input('Command: ')
            if self.command in commands:
                break
    
    def main_loop_saying(self):
        self.words_index = list(range(len(self.sel_cat_words.index)))
        random.shuffle(self.words_index)
        print('\nCommands:\n-next (skip flashcard)\n-tryagain (try again to say correct word)\n-answer (show correct answer)\n-menu (go back to menu)\n-exit (exit program)\n')
        for index in self.words_index:
            print(f"{self.sel_language}: {self.sel_cat_words.at[index, self.sel_language]}")
            FC.say_text(self.sel_voice, self.sel_cat_words.at[index, self.sel_language])
            while True:
                print(f"Say this word in {self.sec_language}")
                with self.mic as source:
                    self.r.adjust_for_ambient_noise(source)
                    self.audio = self.r.listen(source)
                try:
                    self.speech_to_text = self.r.recognize_google(self.audio, language=self.rec_language)
                except:
                    continue
                else:
                    if self.speech_to_text == self.sel_cat_words.at[index, self.sec_language].lower():
                        print(f"{self.sec_language}: {self.speech_to_text}")
                        FC.get_command(['-next', '-menu', '-exit'])
                        if self.command == '-next':
                            break
                        elif self.command == '-menu':
                            FC.menu()
                        elif self.command == '-exit':
                            exit()
                    else:
                        print('Wrong')
                        FC.get_command(['-next', '-tryagain', '-answer', '-menu', '-exit'])
                        if self.command == '-next':
                            break
                        elif self.command == '-tryagain':
                            continue
                        elif self.command == '-answer':
                            print(f"{self.sec_language}: {self.sel_cat_words.at[index, self.sec_language]}")
                            FC.say_text(self.sec_voice, self.sel_cat_words.at[index, self.sec_language])
                        elif self.command == '-menu':
                            FC.menu
                        elif self.command == '-exit':
                            exit()

        FC.end_menu()

    def main_loop_writing(self):
        self.words_index = list(range(len(self.sel_cat_words.index)))
        random.shuffle(self.words_index)
        print('\nCommands:\n-next (skip flashcard)\n-answer (show correct answer)\n-menu (go back to menu)\n-exit (exit program)\n')
        for index in self.words_index:
            print(f"{self.sel_language}: {self.sel_cat_words.at[index, self.sel_language]}")
            while True:
                self.user_input = input(f"{self.sec_language}: ")
                if self.user_input == self.sel_cat_words.at[index, self.sec_language].lower():
                    break
                elif self.user_input == '-next':
                    break
                elif self.user_input == '-answer':
                    print(f"{self.sec_language}: {self.sel_cat_words.at[index, self.sec_language]}")
                    break
                elif self.user_input == '-menu':
                    print('')
                    FC.menu()
                elif self.user_input == '-exit':
                    exit()
                
        FC.end_menu()

    def end_menu(self):
        print('\nEnd of words in this category\nMenu [1]\nExit [2]')
        while True:
            try:
                self.choice = int(input('Choice: '))
            except ValueError:
                continue
            if self.choice == 1:
                print('')
                FC.menu()
            elif self.choice == 2:
                exit()



FC = Flashcards()
FC.menu()



    
