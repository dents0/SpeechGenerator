Speech Generator
================
Licensing information: READ LICENSE
----
Project source can be downloaded from https://github.com/dents0/SpeechGenerator.git
----

Author
------
Deniss Tsokarev

Description
-----------
A speech generator written in Python. 
The script generates a speech of the desired lenght and saves it to a .docx file. 

People say approximately 5 sentences per minute, however, as the speech gets longer, the amount of sentences per minute will decrease. 

The less length is chosen the less repetitive the speech will look.

Requirements
------------
* Python 3.6+
* python-docx module

Here is how to install **python-docx** with **pip** using a command-line interpreter on Windows machines:
```
pip install python-docx
```

How to use
----------
1) Run the file **speech_generator.py** using command-line interpreter.
2) Enter the desired **length** of speech, the **heading** and **name** the file that will store the result.

![screenshot1](https://user-images.githubusercontent.com/28843507/52724663-0e2b9e80-2fb0-11e9-9b1d-7d4cc7fe1543.PNG)

3) The file with generated speech will be created in the same directory where the file speech_generator is.

![screenshot2](https://user-images.githubusercontent.com/28843507/52724778-47640e80-2fb0-11e9-84d4-e1dca5ffa101.PNG)

4) The speech will have the heading you created and will be divided into paragraphs to improve readability.

![screenshot3](https://user-images.githubusercontent.com/28843507/52724784-49c66880-2fb0-11e9-8dd5-fd7dded436c4.PNG)
