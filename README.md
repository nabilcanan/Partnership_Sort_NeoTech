# Partnership_Sort_NeoTech


# Currently, we select the latest Neotech file and then the previous one 

    -We need to fix the issue where we need to check the latest neotechfile 
    
    -Ideally, once the process is completed for the last week file we can bring over the new VLOOKUP'ed columns 
    and bring them over
    
    -I need the week of 8.18 contract to have the Dupes Removed Sheet, with the VLOOKUP columns from the 8.04 
    contract in order to bring them into the next contract being 9.01

    - We need to make it so no matter what the Dupes Removed or in our case Full Fle Without Dupes is created to then use the rest of our functions accordingly
    
    - Currently I can rename the 9.01 first sheet to 'Sheet1' to mock the 9.18 one, then delete the Dupes Removed, once I did this I can run my program on it and I dont need to drop those columns commented out in my code 
      I dont need to drop these anymore because in our latest file those columns aren't there anymore 

    - Now that I have the columns I need in the latest NeoTech file I can choose the latest one to merge the files into

    - Basically all we need is to setup 8.18 the same way 8.04 is setup 

    - Now all we need now that 8.18 is setup is to vlookup over those 4 columns that she has and we can run our program 

    - Scratch this, now 8.18 is updated in my final file.... the 8.18 one

    - We can now run the 9.01 file, and the 8.18 file and we get our desired function, we get all the columns Vlookuped 
    we get our removed from prev file sheet, we get the duplicates removed 

    - Run and double check logic to ensure we are producing the correct data 

    - Need to add the ability to bring over formulas that were implemented into our original excel file 
    
    - Change Full File Without Dupes - Dupes Removed

    - Add a column into the final product called Contract Change that checks the Base Unit Price changes, increaces, decreases, new or no change 