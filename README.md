# Exchange Attachment Grabber

A temporary solution for a hands-off way to grab EDI files from clients exchange servers for importing.

It loops through each mail item in the inbox, checks it is from a whitelisted address, and then checks each attachment in the mail file. 

It will check that the file is named as our importing software expects, and if it is, it downloads the attachment to the specified local folder, and appends the file with the company name and date-time downloaded.

Once done with each mail item, it gets moved to an 'Imported EDI' mail folder so it doesn't get read again.