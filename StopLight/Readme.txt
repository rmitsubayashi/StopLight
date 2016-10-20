**********
StopLight
**********

--Introduction--
This Microsoft Word Add-In is intended to assist foreign language teachers in making tests for class.
With their abundant knowledge, teachers are likely to incorporate high-level words in their tests.
This program allows visual representation of word difficulty levels
by highlighting the words in green, yellow, or red (stop light colors!) according to difficulty level.
Words not in the data set will be highlighted in gray.
The data is read from an Excel file, in which the user can modify before or at runtime.

--Prerequisites/Installation--
You will need Visual Studios installed to run the program.
Please import the project from Github and build the solution on Visual Studios.
You can deploy the project from Visual Studios using their publishing wizard.

--User Story Example--
Sherry is a middle school English teacher in Japan. She makes a quiz for class tomorrow,
but she is worried that she used words they haven't covered yet.
She decides to use StopLight to decide whether the quiz is appropriate for their level.
She reads in an Excel file where the columns are organized by the page numbers on the class's textbook.
For 'none', she sets to range to the page numbers preceding two weeks ago.
For 'green', she sets the range to two weeks ago to today's page number.
For 'yellow', she sets the range to the next page number, hoping some students might have already started studying!
For 'red', she sets the range to every page number beyond the next page.
Sherry was able to identify from these highlights that
-her quiz has a supple amount of target words, indicated by green highlighting
-her quiz contains several high level words, which Sherry will adjust now
-SpotLight is such and awesome program!!

--Limitations--
-Compound words such as "traffic jam" and "merry-go-round" will be counted as multiple words.
-The lemmatizer considers word by word, so semantics are not involved in the identification of a word.
  For example in the sentence 'I saw Betty', 'saw' will be interpreted as 'saw' as in a saw mill, rather than 'see'.
-No functionality to remove/edit words in the Excel file at runtime (for now). Will add if requested.
-Currently, the program supports English classification and English/Japanese UI. Will add more as requested.

--Excel File Formatting--
For the Excel file,
  1. The first row shall consist of unique classification names for the data
     Ex: page number, date, grade
  2. Every column below the first row shall consist of every word that fit in the classification criteria for each column name.
For more information, please see the example Excel file 'example.xlsx'

--Example Excel File--
The example Excel file 'example.xlsx' in the Data folder consists of data provided by http://www.eigo-duke.com.
These are English words learned in Japanese schools in middle and high school, divided by grade of aquisition.
The words 'be', 'a', and 'the' have been added to the list for better initial usability.

--Lemmatization--
For lemmatizing words to make it easier to classify, StopLight uses the open source platform LemmaGen.

--Author--
Ryosuke Mitsubayashi - Github username : rmitsubayashi

--License--
Copyright [2016] [Ryosuke Mitsubayashi]

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License. 