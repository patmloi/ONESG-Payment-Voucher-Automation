# ONESG-Payment-Voucher-Automation
This program aims to streamline ONESG's payment voucher (PV) process, in three parts: 

## 1. Creating a PV from user inputs
Users input information from the images sent. The code then automatically formats the information according to the format specified by ONESG, and enters the formatted information out a copy of the template PV.
   
## 2. Sends an email to the approval staff with the PV and supporting images attached
The code sends an email with a default message and the specified attachments. A custom message can be added to indicate any anomalies in the process. 

## 3. Uploads the supporting images to Dropbox
The code creates a new folder in Dropbox, and uploads the supporting images in it. The PV is not uploaded at this stage, as it would still be pending approval. 
