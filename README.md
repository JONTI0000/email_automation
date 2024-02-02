# --This is a set of instructions on how to set up this automation tool--
NOTE
each batch file can only be accessed by one email at a time


when setting up the file structure for this the following folders should be given as well as all the files should be kept in the
respective folder accordingly

-batches(folder) should contain the batch file in an xlsx format with the naming structure with the following format
batch file 1 would have the name '01.xlsx'

-steps(folder) should contain the necassary 4 steps in txt format these should be the html txt files and each step file can have the 
following placeholder values they are 
	{name}:name of the receiver
	{area}:area of the business
these place holder values can be used anywhere in the template.

-signature(file) should contain the signature in the necassary format in this case there are 2 placeholder values that should be there they are	
	{name}:name of the sender
	{email}:email of the sender

with these following things setup to execute the process double click on the run file
with that the following prompts will be displayed

How many Email addresses: {This asks for the number of email addresses that should be used to send out the emails}

there after it will be asking the neccassary parameters for each email address
Name of Sender: {this should be the name assigned to the email address}

Enter the Step: {this would be the step that should be going out} 
NOTE:in order to this to work step should be in the steps folder and step input should be the file name which as an example for the first step would be '01'

Enter the Subject: {this would be to enter the neccassary subject for each step}
NOTE:if in the subject the area should be mentioned {} use this as the placeholder it will automatically be replaced with the area of the business

Enter batch no: {this would be to enter the batch no}
NOTE:in order to this to work batch no file should be in the steps folder and batch no input should be the file name which as an example for the first batch would be '01'
