# largeTransactionJsonBuilder
Install steps  

1. Pull the code into local repo  
2. `python -m venv env`  
3. `source env/scripts/activate`  
4. `pip install -r requirements.txt`  


# Preparing the sheet  
It's important to prepare the sheet carefully, otherwise the json will not be correct.  
The first row is the internal value.  
Second row the external value.  
![image](https://user-images.githubusercontent.com/74533741/189054041-2b592ba1-4a3e-4c1d-ae1c-b9dfcb33175e.png)

When ready, enter in terminal `python main.py`  
Enter the filename for the target sheet, important to add the .xlsx extension  
![image](https://user-images.githubusercontent.com/74533741/189053390-f0ba0788-9ede-497c-aa77-fff426a53892.png)

![image](https://user-images.githubusercontent.com/74533741/189054305-62527b57-e983-45d6-a93f-783a8bcaa7fc.png)

If everything was prepared correctly, you should see the json object in the last column  
![image](https://user-images.githubusercontent.com/74533741/189054606-4723bc5d-6d8c-4970-8673-c9e7daf3693c.png)

![image](https://user-images.githubusercontent.com/74533741/189054525-872fb05e-6529-4ec2-aad2-5ba231f10282.png)

