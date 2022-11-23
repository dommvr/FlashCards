# FlashCards

Flashcard app using Excel database  
By using it you can learn new words by listening, reading, saying and writing them

# Setup
 - You can use any words database, just put it in project folder and name it `words.xlsx` 
    or change database path:
``` 
self.df_words = pd.read_excel('words.xlsx')
```
- Any new Excel database should have same format as default one i.e. 2 columns have 
  languages names and all words with it's category in 3rd column
  ![Excel database](https://imgur.com/a/iMFALEM)
- By default app supports only English and Polish. In order to use another language you need 
  to [download language pack for speech](https://support.microsoft.com/en-us/windows/download-language-pack-for-speech-24d06ef3-ca09-ddcc-70a0-63606fd16394)
  (if you don't have it already installed) 
- Next you need to add new voice and [language tag](https://www.techonthenet.com/js/language_tags.php) to dictionary
``` 
self.voices = {
    "English": ["HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0", "en-US"],
    "Polish": ["HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_PL-PL_PAULINA_11.0", "pl-PL"]
}
```
# Usage
 - Once you run `main.py` you can select languages, form of learning and words category
