# PARSING-MINISTRY-OF-JUSTICE-OF-THE-SLOVAK-REPUBLIC-COMMERCIAL-REGISTER-DATABASE

The code is demonstrated in the file `sldb.py`.

Parsing data from www.orsr.sk/default.asp.

For data parsing, a method of searching by organization name was chosen, using this link: https://www.orsr.sk/search_subjekt.asp.

The biggest problem with parsing this resource was that the search does not return more than 500 results, and it was necessary to devise a method to reduce the number of results to 500 or less. Searching by organization name simplifies this task by allowing the selection of the organization form and court.

The code outputs two documents in xlsx format: initial_results.xlsx is needed for debugging the code, while final_results.xlsx is the final result.

Python libraries used in the code:

Requests - for making HTTP requests;

BeautifulSoup - for parsing HTML documents;

Xlsxwriter - for creating Excel files;

Openpyxl - for editing Excel files;

ThreadPoolExecutor - for more efficient and faster code execution;

Time - for measuring the execution time of the code.


The methodology is demonstrated in the form of a process map (Process Map.pdf).

To speed up the execution of the code, multithreading technology was used.
