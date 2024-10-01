// uploadTemplate.js
const mongoose = require('mongoose');
const Grid = require('gridfs-stream');
const fs = require('fs');

mongoose.connect('mongodb://192.168.50.57:27017/local', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const conn = mongoose.connection;

conn.once('open', () => {
  const gfs = Grid(conn.db, mongoose.mongo);

  const writeStream = gfs.createWriteStream({
    filename: 'OP###_Client_GridOS-DF_SER_v0.0.docx',
    content_type:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    root: 'std_contents', // Specify the collection name
  });

  fs.createReadStream(
    'C:/Users/rikog/Box/_CSE_PROPOSALS/ARTEFACTS/GOS_gridos/std_contents/OP###_Client_GridOS-DF_SER_v0.0.docx'
  ).pipe(writeStream);

  writeStream.on('close', (file) => {
    console.log('Template uploaded successfully:', file);
    process.exit(0);
  });

  writeStream.on('error', (err) => {
    console.error('Error uploading template:', err);
    process.exit(1);
  });
});
