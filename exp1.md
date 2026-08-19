```python
!pip install openpyxl
```

    Collecting openpyxl
      Downloading openpyxl-3.1.5-py2.py3-none-any.whl.metadata (2.5 kB)
    Collecting et-xmlfile (from openpyxl)
      Downloading et_xmlfile-2.0.0-py3-none-any.whl.metadata (2.7 kB)
    Downloading openpyxl-3.1.5-py2.py3-none-any.whl (250 kB)
    Downloading et_xmlfile-2.0.0-py3-none-any.whl (18 kB)
    Installing collected packages: et-xmlfile, openpyxl
    
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       -------------------- ------------------- 1/2 [openpyxl]
       ---------------------------------------- 2/2 [openpyxl]
    
    Successfully installed et-xmlfile-2.0.0 openpyxl-3.1.5
    

    
    [notice] A new release of pip is available: 25.2 -> 26.2.1
    [notice] To update, run: python.exe -m pip install --upgrade pip
    


```python
import pandas as pd
import openpyxl

```


```python
df=pd.read_csv('car_dataset_40x20.csv')
print(df.head())

```

      Car_ID     Brand   Model  Year Fuel_Type Transmission  Engine_CC  \
    0   C001  Mahindra  XUV300  2015       CNG       Manual       1199   
    1   C002    Maruti  Brezza  2021    Petrol       Manual       1248   
    2   C003  Mahindra    Thar  2018    Petrol       Manual        998   
    3   C004     Honda   Amaze  2018    Hybrid       Manual       1197   
    4   C005     Honda    City  2022    Diesel    Automatic       1499   
    
       Mileage_kmpl  Price_Lakh  Owner_Count  Kilometers_Driven   Color Body_Type  \
    0          21.9       11.94            2             145442  Silver       SUV   
    1          18.8        7.96            1              66097  Silver       SUV   
    2          22.2       11.39            1             111848   Black       SUV   
    3          17.1       10.21            1              82848   Black     Sedan   
    4          13.1       12.00            1              37049   Brown     Sedan   
    
       Seating_Capacity        City       Seller_Type Insurance_Valid  \
    0                 5   Hyderabad  Certified Dealer             Yes   
    1                 5  Vijayawada  Certified Dealer             Yes   
    2                 4  Vijayawada        Individual              No   
    3                 5   Hyderabad        Individual              No   
    4                 5  Vijayawada  Certified Dealer              No   
    
      Accident_History  Service_Cost_Rs  Resale_Value_Lakh  
    0               No             8168               2.78  
    1               No             6280               4.20  
    2               No             8007               4.10  
    3            Minor             7750               3.67  
    4               No            13349               7.20  
    


```python
df.shape
```




    (40, 20)




```python

df.columns

```




    Index(['Car_ID', 'Brand', 'Model', 'Year', 'Fuel_Type', 'Transmission',
           'Engine_CC', 'Mileage_kmpl', 'Price_Lakh', 'Owner_Count',
           'Kilometers_Driven', 'Color', 'Body_Type', 'Seating_Capacity', 'City',
           'Seller_Type', 'Insurance_Valid', 'Accident_History', 'Service_Cost_Rs',
           'Resale_Value_Lakh'],
          dtype='str')



# writing dat into csv


```python
Student_data={
    "Roll_no":[10,11,12,13],
    "name":["Ganesh","mani","nithin","Chandhu"],
    "Dept":["CSE","IT","ECE","CSE"],
}
df=pd.DataFrame(Student_data)
df
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Roll_no</th>
      <th>name</th>
      <th>Dept</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>10</td>
      <td>Ganesh</td>
      <td>CSE</td>
    </tr>
    <tr>
      <th>1</th>
      <td>11</td>
      <td>mani</td>
      <td>IT</td>
    </tr>
    <tr>
      <th>2</th>
      <td>12</td>
      <td>nithin</td>
      <td>ECE</td>
    </tr>
    <tr>
      <th>3</th>
      <td>13</td>
      <td>Chandhu</td>
      <td>CSE</td>
    </tr>
  </tbody>
</table>
</div>




```python
df.to_csv("Student_output.txt")
print("Succes")
```

    Succes
    


```python
df
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Roll_no</th>
      <th>name</th>
      <th>Dept</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>10</td>
      <td>Ganesh</td>
      <td>CSE</td>
    </tr>
    <tr>
      <th>1</th>
      <td>11</td>
      <td>mani</td>
      <td>IT</td>
    </tr>
    <tr>
      <th>2</th>
      <td>12</td>
      <td>nithin</td>
      <td>ECE</td>
    </tr>
    <tr>
      <th>3</th>
      <td>13</td>
      <td>Chandhu</td>
      <td>CSE</td>
    </tr>
  </tbody>
</table>
</div>




```python
nf=pd.read_csv("Student_output.txt")
nf
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Unnamed: 0</th>
      <th>Roll_no</th>
      <th>name</th>
      <th>Dept</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>0</td>
      <td>10</td>
      <td>Ganesh</td>
      <td>CSE</td>
    </tr>
    <tr>
      <th>1</th>
      <td>1</td>
      <td>11</td>
      <td>mani</td>
      <td>IT</td>
    </tr>
    <tr>
      <th>2</th>
      <td>2</td>
      <td>12</td>
      <td>nithin</td>
      <td>ECE</td>
    </tr>
    <tr>
      <th>3</th>
      <td>3</td>
      <td>13</td>
      <td>Chandhu</td>
      <td>CSE</td>
    </tr>
  </tbody>
</table>
</div>




```python
df.columns
```




    Index(['Roll_no', 'name', 'Dept'], dtype='str')



# JSON


```python
Student_data={
    "Roll_no":[10,11,12,13],
    "name":["Ganesh","mani","nithin","Chandhu"],
    "Dept":["CSE","IT","ECE","CSE"],
}
df=pd.DataFrame(Student_data)
df.to_json("Students.json",orient='records',indent=4)
print("JSon file created")
```

    JSon file created
    


```python
df=pd.read_json('Students.json')
print("Student Data")
df
```

    Student Data
    




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Roll_no</th>
      <th>name</th>
      <th>Dept</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>10</td>
      <td>Ganesh</td>
      <td>CSE</td>
    </tr>
    <tr>
      <th>1</th>
      <td>11</td>
      <td>mani</td>
      <td>IT</td>
    </tr>
    <tr>
      <th>2</th>
      <td>12</td>
      <td>nithin</td>
      <td>ECE</td>
    </tr>
    <tr>
      <th>3</th>
      <td>13</td>
      <td>Chandhu</td>
      <td>CSE</td>
    </tr>
  </tbody>
</table>
</div>



# Excel


```python
Student_data={
    "Roll_no":[10,11,12,13],
    "name":["Ganesh","mani","nithin","Chandhu"],
    "Dept":["CSE","IT","ECE","CSE"],
}
df=pd.DataFrame(Student_data)
df.to_excel("Student.xlsx",sheet_name="Student details",index=False)
print("Data written succesfully in excel")
```

    Data written succesfully in excel
    


```python
df=pd.read_excel("Student.xlsx",sheet_name="Student details")
print("Data from excel file")
df
```

    Data from excel file
    




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Roll_no</th>
      <th>name</th>
      <th>Dept</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>10</td>
      <td>Ganesh</td>
      <td>CSE</td>
    </tr>
    <tr>
      <th>1</th>
      <td>11</td>
      <td>mani</td>
      <td>IT</td>
    </tr>
    <tr>
      <th>2</th>
      <td>12</td>
      <td>nithin</td>
      <td>ECE</td>
    </tr>
    <tr>
      <th>3</th>
      <td>13</td>
      <td>Chandhu</td>
      <td>CSE</td>
    </tr>
  </tbody>
</table>
</div>




```python

```


```python

```
