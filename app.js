const express = require('express');
const app = express();
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const dbo = require('./db');
const ObjectID = dbo.ObjectID;
const collectionName='details';
//const YourModel=require('./schema')


app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

async function readExcelAndInsertData() {
    try {
      const database = await dbo.getDatabase();
      const collection = database.collection(collectionName);
  
      // Read data from the Excel sheet
      const workbook = xlsx.readFile('Book 3.xlsx');
      const sheetName = workbook.SheetNames[0];
      const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
  
      // Filter out rows with empty values
      const filteredData = sheetData.filter(item => Object.values(item).some(val => val !== null && val !== ''));
  
      // Insert data into MongoDB
      const result = await collection.insertMany(filteredData);
      console.log(`${result.insertedCount} documents inserted successfully.`);
    } catch (err) {
      console.error('Error:', err);
    }
  }
  
  readExcelAndInsertData()

// CRUD operations

app.get('/read', async (req, res) => {
    try {
        const database = await dbo.getDatabase();
        const collection = database.collection(collectionName);
    
        // Fetch data from MongoDB
        const items = await collection.find({}, { projection: { _id: 0 } }).toArray(); 
        const filteredItems = items.filter(item => Object.values(item).some(val => val !== null && val !== ''));   
        // Create a new workbook and worksheet
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(filteredItems);
    
        // Add the worksheet to the workbook
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Data');
    
        // Save the data to a new Excel file and send it in the response
        xlsx.writeFile(workbook, 'data.xlsx');
        res.download('data.xlsx', 'data.xlsx'); // Provide the file for download
      } catch (err) {
        console.error('Error:', err);
        res.status(500).json({ error: 'An error occurred.' });
      }
    
});

app.post('/insert_details', async (req, res) => {
  try {
    const database = await dbo.getDatabase();
    const collection = database.collection(collectionName);
    const newItem = req.body;
    await collection.insertOne(newItem);
    res.status(201).json(newItem);
  } 
  catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'An error occurred.' });
  }
});

app.put('/update_detail/:id', async (req, res) => {
  try {
    const database = await dbo.getDatabase();
    const collection = database.collection(collectionName);
    /*const updatedItem = req.body;
    let details = { id: req.body.id, name: req.body.name , age:req.body.age, email: req.body.email, phone:req.body.phone,postalCode:req.body.postalCode };*/
    const itemId = req.params.id;
    const objectId = new ObjectID(itemId);
    const unsetFields = req.body;
    await collection.updateOne({ _id: objectId}, { $set: unsetFields });    
    res.json({ message: 'Data updated successfully.' });
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'An error occurred.' });
  }
});

app.delete('/delete_detail/:id', async (req, res) => {
  try {
    const database = await dbo.getDatabase();
    const collection = database.collection(collectionName);
    const itemId = req.params.id;
    const objectId = new ObjectID(itemId);
    const result=await collection.deleteOne({ _id: objectId });
    if (result.deletedCount === 1) {
        res.json({ message: 'Data deleted successfully.' });
      } else {
        res.status(404).json({ error: 'Data not found.' });
      }
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'An error occurred.' });
  }
});

const port = 3000;
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
