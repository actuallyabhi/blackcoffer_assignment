# Steps to run

1. Change the directory to the project directory
2. Install the requirements
    pip install -r requirements.txt
3. Run the project
    python main.py

# Project Structure
    
    ``` 
    ├── README.md
    ├── requirements.txt
    ├── Input.xlsx(contains the input data)
    ├── main.py
    ├── WordLists
    │   ├── MasterDictionary
    │   │   ├── negative-words.txt
    │   │   ├── positive-words.txt
    │   ├── StopWords
    │   │   ├── StopWords_Auditor.txt
    │   │   ├── StopWords_Currencies.txt
    │   │   ├── StopWords_DatesandNumbers.txt
    │   │   ├── StopWords_Generic.txt
    │   │   ├── StopWords_GenericLong.txt
    │   │   ├── StopWords_Geographic.txt
    │   │   ├── StopWords_Names.txt
    ├── Output.xlsx(contains the output data)
    |-- TextFiles(contains the text files)
    |   ├── TextFile1.txt
        ...
    |   ├── TextFileN.txt

    ```
