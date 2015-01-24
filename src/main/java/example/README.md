



The Objective of PSTExtractor to provide an example to Extract messages from the PST files and its attachments.
This class will print the result of extracting messages, TO, FROM, CC, subject and body from messages to console,
the objective of this is to simplify the indexing process, so if you want to index the content of a single pst file
 should pipe the console to a file and index it.
Some limitations of this code:
- Doesn't set up the permissions of the Attachments as the pst file has, it creates with the permissions of the process.
 <b>But in linux you can use setfacl/set-acl in order to do this change.</b>
- Doesn't prepare a specific format for further indexing with search engines like GSA, but you can use the <b>console ouput 
 to generate a valid format.</b>
- Doesn't extract the content of the attachments, meaning the information that is inside the file attached to the email message.
- This is not a proper crawler, so won't get a list of PST inside a folder and transform all of them, but here you have 
 an example that will work on many of the UNIX like systems: 
 <b>$find /Users/iolalla/workspace/ -name "*.pst" -exec java -cp java-libpst-0.8.1.jar:javax.mail-1.5.2.jar:. example.PSTExtractor {} \;</b>
 
Finally, in order to run this class here you have an example you can use, from target dir :
$java -cp java-libpst-0.8.1.jar:javax.mail-1.5.2.jar:. example.PSTExtractor /Users/iolalla/workspace/java-libpst/YOUR.pst
