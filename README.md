# Wine_List_Parser
A python script that reads in a number of Wine Lists in the form of word documents and outputs a single document with all the wines that were on any of the lists

Backstory: 
I work at a small Wine bar where we are constantly changing the wines on our menu. The list is around 14 pages long and all the wines have a short paragraph of tasting notes. Whenever we changed the list we would just duplicate the most recent wine list and make changes. Usually copy and pasting from previous lists. Then save the old list to an archive folder. As we amassed more and more lists in the archive folder, it became slower and more tedious to search through old menus for the wine and tasting notes that we were looking for. 

Solution: 
I saw this as an opportunity to use my programming skills and apply it to a real world inefficiency.  My plan was the write a script that I could feed word documents, then it would output a single word document, with all the wines that we have ever had on the list, sorted alphabetically and correctly formatted. 
To further make it more convenient for work would be to run it when ever a new wine list is saved to the folder. 

Program: 
The program mainly uses python-docx to parse the document into python and then also to output the single list. It first converts all the lists to txt files to all python quickly analyse them. The get_wines() function then goes through each of the lists to find the wines and tasting notes. The structure of the wines and notes in the menu are reasonably consistent. 

“Year” “Name” | “Region” | Price Bottle | Price Glass 
Notes…………

However, some of them don’t have a region, and we don’t sell them all by the glass, so sometimes there isn’t a glass price. 
The function read each line and sees if it matches the pattern. While I would have liked to do this with regular expressions, there just was not enough consistency in the formatting to do it. 

Overall,, I am super proud of my work on this. This is my first 'real' project that i have worked to completion. It has been a lot of fun, and I have learnt a lot from it. 
