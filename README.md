# ERACWonkinator
ERAC Called this the Wonkinator. :-)

The problem:  Insurance Claim supporting docs need to be sent to 3rd parties via physical mail.  Thousannds, daily.  The print room at ERAC has 3 paths, depending on the # of sheets of paper, but the problem is they have a handfull of digital files, and no idea how many pages it will print to, and all these files are different file types.

This was one of the first projects I worked on at Adlib, I coded it one night in a hotel room in Germany where I rushed to get it done since my friend and I were anxious to get out to Karaoke at an Irish pub.  In Berlin.

The process is a bit more complex than just normalizing documents to PDF and counting their pages, we also have to apply barcodes to each page so that the sorting and stuffing machines will know when to grab a new envelope, and these barcodes must be in series, one series per line, with 3 lines of envelope stuffing.

Since Adlib Express 4.x is no longer available for sale, and ERAC has migrated this logic long ago into a more integrated process, I'm putting these files here for reference.

One project contains the code to monitor the folder for folders, the other project manages all of the logic after normalizing and combining folder contents to a single PDF File, identifying the number of sheets, applying the barcodes and placing the output in the right folder to initiate the physical process (The actual printing & envelope stuffing)
