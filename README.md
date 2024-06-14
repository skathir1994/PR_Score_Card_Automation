# PR_Score_Card_Automation
NVA Reduction on Score Card Work.

Problem Statement: PR team prepared the PR score card every week as part of the DPMU reporting process where valid rejections are aligned with the root cause and categorized into controllable and uncontrollable defects.  This process currently is an entirely manual process. The score card contains 35 columns & n-number of rows for each rejection tracked YTD. The first 21 columns in the PR Score card are from PR team data base. The remaining 14 columns have to be filled in manually as they are dependent on other data sources. 
(Ex: Rejection Source Colum is dependence to (PA/DSM/PDM/Andon) team member id). Current time taken to complete this activity is 1 productive hour/week.    

Proposed Solutions: For this problem, I have come up with a pythod code-based solution. Using this predefined encoded script, 14 columns get updated in an automated manner. Process to be followed is load the input file from PR DB (First 22 columns) and run the code which auto-populates the remaining fields in the table and generates the final output, the PR score card, containing a total of 35 columns. 

Listed below are encoded logics.
![image](https://github.com/skathir1994/PR_Score_Card_Automation/assets/66460217/8d46d3c1-e6d7-40f0-9765-65f204e621b0)


Impact : The entire is completed within  2 mints, with a saving of 50 Mins each week.  
