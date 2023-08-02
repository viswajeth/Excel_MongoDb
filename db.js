const mongodb= require('mongodb');
const { MongoClient } = require('mongodb');
const ObjectID= mongodb.ObjectId;


const mongoURI = 'mongodb+srv://tsrviswajeth:viswajeth@excel.4nctr3z.mongodb.net/?retryWrites=true&w=majority';
const dbName = 'Excel';
//const collectionName = 'details'


async function getDatabase() {
    const client = await MongoClient.connect(mongoURI, { useUnifiedTopology: true });
    
    if(!dbName){
        console.log('Database not connected');
    }
    return client.db(dbName);
  }

  module.exports={
    getDatabase,
    ObjectID
  }