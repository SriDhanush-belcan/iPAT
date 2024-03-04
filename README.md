# iPAT
Initial Package Automation Tool

                                             Automation Project
Requirements:
Total input PPT files = 3
-	SAP
-	Vendor
-	Automation( this is expected output file, the titles should be copied from this)
  
Total output files = 1
-	xxxx.pptx
  
Total no.of titles = 
1.	QN RESOLUTION DATA 
- Reference object( copy from SAP)
- Here there will be a text box where user input data will be printed here, QN number will be pasted here and will be applied to other slides where ever QN number is mentioned.(AutomationPPT-outputfile expected)
  
2.	 Defects & long text – Line item ‘X’
-	Here X represents no.of line items, depending on no.of line items the slides should be created, this ‘X’ will be given in the first slide creation only. 
-	Example: There will be a dialogue box displayed in the beginning where users gives input number for the line items in the text box along with the other things mentioned in the given Ref.PPT(automation.ppt), based on that the PPT file will be created with those many slides
  
3.	Defect MQI Location 
4.	Disposition & Approvals
5.	Design Assessment
6.	Engine Cross Section(*)
7.	Interfacing Parts 
8.	Interfacing Parts (3D View)[Manual]
9.	Vendor Information(This slide should be copied Separate PPT file)
10.	Engine Manual
11.	Root Cause & Corrective Action
12.	Prior History by S/N
13.	Prior History by MQI or Damage Code
    
GREEN TEXT REPRESENTS MANUAL INTERVENTION*

CONCLUSION:

These titles should be created in the output file based on the user input values the slides should be created i.e, Line item number. 
Input file name – SAP.pptx , Vendor.pptx, Automation.pptx
Output file name – XXX.pptx, Automation.ppt (reference given to developer on this name)

