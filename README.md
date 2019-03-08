# utl-importing-all-emails-in-a-MS-outlook-folder-into-SAS
Importing all emails in a MS outlook folder into SAS 
    Importing all emails in a MS outlook folder into SAS                                                                            
                                                                                                                                    
    If you happen to have                                                                                                           
                                                                                                                                    
      Outlook 2010 64bit                                                                                                            
      Ein  7 64bit                                                                                                                  
      Classic Workstaion SAS 94M2 64bit and Access to PC-Files                                                                      
                                                                                                                                    
      Microsoft is a moving target so this may only work on July 4th 2010?                                                          
                                                                                                                                    
      There are probably python modules to read outlook folders (pyOutlook?)                                                        
      Maybe R and Perl too.                                                                                                         
    *_                   _                                                                                                          
    (_)_ __  _ __  _   _| |_                                                                                                        
    | | '_ \| '_ \| | | | __|                                                                                                       
    | | | | | |_) | |_| | |_                                                                                                        
    |_|_| |_| .__/ \__,_|\__|                                                                                                       
            |_|                                                                                                                     
    ;                                                                                                                               
                                                                                                                                    
    Create some emails and save the emails in an outlook folder.                                                                    
    I am using the drafts folder.  Because I don't want Outlook connected to the web.                                               
                                                                                                                                    
    Outlook Drafts folder                                                                                                           
                                                                                                                                    
    --------------------                                                                                                            
                                                                                                                                    
    Subject: March Sales                                                                                                            
                                                                                                                                    
                                                                                                                                    
    To JohnDoe@doe.con                                                                                                              
    CC JaneDoe@doe.con                                                                                                              
                                                                                                                                    
                                                                                                                                    
     03MAR2019 Sales 190875                                                                                                         
     04MAR2019 Sales 290875                                                                                                         
     05MAR2019 Sales 390875                                                                                                         
     06MAR2019 Sales 490875                                                                                                         
     07MAR2019 Sales 590875                                                                                                         
     08MAR2019 Sales 690875                                                                                                         
                                                                                                                                    
    ----------------------                                                                                                          
                                                                                                                                    
    Subject: April Sales                                                                                                            
                                                                                                                                    
                                                                                                                                    
    To JohnDoe@doe.con                                                                                                              
    CC Jane@doe.con                                                                                                                 
                                                                                                                                    
                                                                                                                                    
     03APR2019 Sales 290875                                                                                                         
     04APR2019 Sales 290875                                                                                                         
     05APR2019 Sales 390875                                                                                                         
     06APR2019 Sales 490875                                                                                                         
     07APR2019 Sales 590875                                                                                                         
     08APR2019 Sales 790875                                                                                                         
                                                                                                                                    
    *            _               _                                                                                                  
      ___  _   _| |_ _ __  _   _| |_                                                                                                
     / _ \| | | | __| '_ \| | | | __|                                                                                               
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                               
                    |_|                                                                                                             
    ;                                                                                                                               
                                                                                                                                    
       1. MS ACCESS table with all emaile in folder drafts                                                                          
       2. Parsed emails in a SAS table                                                                                              
                                                                                                                                    
                                                                                                                                    
    ---------------------------------------------------                                                                             
    1. MS ACCESS table with all emaile in folder drafts                                                                             
    ---------------------------------------------------  
    
     Automatically copy all emails to anaccess table drafts                  
                                                                            
      a. open a blank database in ms access                                 
      b. click on tab external data                                         
      c. click on "more" just below XML file                                
      d. click on outlook folder "drafts"                                   
      e. click on  folder "drafts" > highlight "drafts" > next "ok"         
      f. save database in d:/mdb/drafts.accdb                               
                                                                            
                                                                                                                                                                                                                                                                     
    d:/mdb/drafts.accdb                                                                                                             
                                                                                                                                    
                                                                                                                                    
     MS Access Table  DRAFTS                                                                                                        
                                                                                                                                    
      Columns                                                                                                                       
                                                                                                                                    
         Name                   Type                 Size                                                                           
         ----------------       --------------       ------                                                                         
                                                                                                                                    
         ID                     Long Integer            4                                                                           
                                                                                                                                    
         Importance             Long Integer            4                                                                           
         Icon                   Text                  255                                                                           
         Priority               Long Integer            4                                                                           
                                                                                                                                    
         Subject                Text                  255                                                                           
         From                   Text                  255                                                                           
         Message To Me          Yes/No                  1                                                                           
                                                                                                                                    
         Message CC to Me       Yes/No                  1                                                                           
         Sender Name            Text                  255                                                                           
         CC                     Text                  255                                                                           
                                                                                                                                    
         To                     Text                  255                                                                           
         Received               Date/Time               8                                                                           
         Message Size           Long Integer            4                                                                           
                                                                                                                                    
         Contents               Memo                    -                                                                           
         Created                Date/Time               8                                                                           
         Modified               Date/Time               8                                                                           
                                                                                                                                    
         Subject Prefix         Text                  255                                                                           
         Has Attachments        Yes/No                  1                                                                           
         Normalized Subject     Text                  255                                                                           
                                                                                                                                    
         Object Type            Long Integer            4                                                                           
         Content Unread         Long Integer            4                                                                           
                                                                                                                                    
                                                                                                                                    
    -------------------------------                                                                                                 
    2. Parsed emails in a SAS table                                                                                                 
    -------------------------------                                                                                                 
                                                                                                                                    
    WORK DRAFTSNRM total obs=12                                                                                                     
                                                                                                                                    
    Obs      SUBJECT             TO                 CC           SALESDATE   SALESAMT         RECEIVED              MODIFIED        
                                                                                                                                    
      1    March Sales    JohnDoe@doe; com    JaneDoe@doe.com    03MAR2019   $190,875    08MAR2019:08:31:00    08MAR2019:08:29:42   
      2    March Sales    JohnDoe@doe; com    JaneDoe@doe.com    04MAR2019   $290,875    08MAR2019:08:31:00    08MAR2019:08:29:42   
      3    March Sales    JohnDoe@doe; com    JaneDoe@doe.com    05MAR2019   $390,875    08MAR2019:08:31:00    08MAR2019:08:29:42   
      4    March Sales    JohnDoe@doe; com    JaneDoe@doe.com    06MAR2019   $490,875    08MAR2019:08:31:00    08MAR2019:08:29:42   
      5    March Sales    JohnDoe@doe; com    JaneDoe@doe.com    07MAR2019   $590,875    08MAR2019:08:31:00    08MAR2019:08:29:42   
      6    March Sales    JohnDoe@doe; com    JaneDoe@doe.com    08MAR2019   $690,875    08MAR2019:08:31:00    08MAR2019:08:29:42   
                                                                                                                                    
      7    April Sales    JohnDoe@doe.con     JaneDoe@doe.con    03APR2019   $290,875    08MAR2019:08:57:16    08MAR2019:08:56:15   
      8    April Sales    JohnDoe@doe.con     JaneDoe@doe.con    04APR2019   $290,875    08MAR2019:08:57:16    08MAR2019:08:56:15   
      9    April Sales    JohnDoe@doe.con     JaneDoe@doe.con    05APR2019   $390,875    08MAR2019:08:57:16    08MAR2019:08:56:15   
     10    April Sales    JohnDoe@doe.con     JaneDoe@doe.con    06APR2019   $490,875    08MAR2019:08:57:16    08MAR2019:08:56:15   
     11    April Sales    JohnDoe@doe.con     JaneDoe@doe.con    07APR2019   $590,875    08MAR2019:08:57:16    08MAR2019:08:56:15   
     12    April Sales    JohnDoe@doe.con     JaneDoe@doe.con    08APR2019   $790,875    08MAR2019:08:57:16    08MAR2019:08:56:15   
                                                                                                                                    
                                                                                                                                    
    *          _       _   _                                                                                                        
     ___  ___ | |_   _| |_(_) ___  _ __                                                                                             
    / __|/ _ \| | | | | __| |/ _ \| '_ \                                                                                            
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                                                           
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                                           
                                                                                                                                    
    ;                                                                                                                               
                                                                                                                                    
    libname mdb access "d:/mdb/drafts.accdb";                                                                                       
                                                                                                                                    
    data draftsNrm ;                                                                                                                
      retain subject to cc from salesdate salesamt received modified ;                                                              
      keep subject to cc salesdate salesamt received modified;                                                                      
      set mdb.drafts;                                                                                                               
      format created modified received datetime25. salesdate date9. salesamt dollar12.;                                             
      do idx=1 by 1 to (countw(contents,"0D0A"x));                                                                                  
          sales=left(scan(contents,idx,'0D0A'x));                                                                                   
          salesamt=input(scan(sales,3),best32.);                                                                                    
          salesdate=input(scan(sales,1,),date9.);                                                                                   
          if sales ne "" then output;                                                                                               
      end;                                                                                                                          
      drop contents idx;                                                                                                            
    run;quit;                                                                                                                       
                                                                                                                                    
                                                                                                                                    
                                                                                                                                    
                                                                                                                                    
                                                                                                                                    
                                                                                                                                    
