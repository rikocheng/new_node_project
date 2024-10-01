const { MongoClient, GridFSBucket } = require('mongodb');
const fs = require('fs');
const path = require('path');

const uri = 'mongodb://localhost:27017/local'; // Your MongoDB URI
const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });

async function uploadFile() {
  try {
    await client.connect();
    const db = client.db('local');
    const bucket = new GridFSBucket(db, { bucketName: 'excel_files' });

    const filePath = path.join(__dirname, 'Client_Resource_Allocation.xlsx'); // Path to your Excel file
    const fileStream = fs.createReadStream(filePath);

    const uploadStream = bucket.openUploadStream('Client_Resource_Allocation.xlsx'); // Name of the file in GridFS

    fileStream.pipe(uploadStream)
      .on('error', (err) => {
        console.error('Error uploading file:', err);
      })
      .on('finish', () => {
        console.log('File uploaded successfully!');
        client.close(); // Close the connection only after the upload completes
      });

  } catch (err) {
    console.error('Error connecting to MongoDB:', err);
  }
}

uploadFile();
