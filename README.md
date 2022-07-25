# excel-from-pojos
This tool converts a list of Plain Old Java Objects (POJOs) to an excel worksheet using reflection, containing all the data from the list of objects.

I opted for reflection because it made the solution more generic to many different object types with completely different attributes.
The idea behind this tool is that it scans the contained class from the list and then for each getter method, it creates a header cell with the attribute name and then using reflection the tool invokes the getter methods from the object to fill the rest of the cells.

Example: 


![image](https://user-images.githubusercontent.com/18034298/180833140-2f57318d-6f89-4af0-bffb-2c8054b5f248.png)
 
![image](https://user-images.githubusercontent.com/18034298/180833232-6bf471f6-52a1-4f0e-a7a6-41ee3e1c49a8.png)

![image](https://user-images.githubusercontent.com/18034298/180833356-a22be94c-efd2-48d7-938b-6450052f067e.png)
