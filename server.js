// server.js
const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const bcrypt = require('bcrypt');
const { GridFSBucket } = require('mongodb'); // Removed MongoClient as it's not used
const app = express();
const jwt = require('jsonwebtoken');
const fs = require('fs');
const multer = require('multer');
const { Readable } = require('stream');
require('dotenv').config();

const storage = multer.memoryStorage();
const upload = multer({ storage });

let wordGridfsBucket;
let excelGridfsBucket;
let stdContentsBucket;

app.use(express.json());




const jwtSecretKey = 'k3G/uwwLInVPjv+IYafV4lDep12NQlk21LfdMSVe4os='; // Use the same secret as in your ONLYOFFICE configuration

// Function to generate a unique document key
function generateDocumentKey() {
  return `doc_${Math.random().toString(36).substring(2, 15)}`; // Generate a random key for the document
}

// Middleware
app.use(bodyParser.json());
app.use(cors({
  origin: 'http://192.168.50.57:8082',
  methods: 'GET,POST,PUT,DELETE',
  credentials: true,
}));

app.use('/web-apps', express.static('/var/www/onlyoffice/documentserver/web-apps', {
  setHeaders: (res, path) => {
    if (path.endsWith('.js')) {
      res.setHeader('Content-Type', 'application/javascript');
    }
  }
}));



// MongoDB connection
const mongoURI = 'mongodb://192.168.50.57:27017/local';
mongoose.connect(mongoURI);

mongoose.connection.on('connected', () => {
  console.log('Connected to MongoDB');
});

mongoose.connection.on('error', (err) => {
  console.log('Error connecting to MongoDB:', err);
});

// Define schemas and models
const userSchema = new mongoose.Schema({
  username: { type: String, unique: true },
  password: String,
});
const User = mongoose.model('User', userSchema);

const userLogsSchema = new mongoose.Schema({
  username: String,
  loginTime: Date,
  action: String, // e.g., "login" or "logout"
});
const UserLogs = mongoose.model('userLogs', userLogsSchema);

// Determine online status based on recent activity
const TIMEOUT_DURATION = 5 * 60 * 1000; // 5 minutes

const dataSchema = new mongoose.Schema({
  clientName: String,
  dataflowEndpoint: String,
  customerApplicationName: String,
  deliveryTimeline: String,
  productionClusterInitial: String,
  productionClusterName: String,
  qualityAndAssuranceClusterInitial: String,
  qualityAndAssuranceClusterName: String,
  customerSolutionName: String,
  maximumLatency: String,
  artifactoryClusterInitial: String,
  artifactoryClusterName: String,
  developmentClusterInitial: String,
  developmentClusterName: String,
  descriptionOfDataflow: String,
  descriptionOfConnectivityBetweenClusterAndLegacyPlatform: String,
  // Add the rest of your data properties...
});
const Data = mongoose.model('Data', dataSchema);

// Initialize GridFSBucket instances
mongoose.connection.once('open', () => {
  const db = mongoose.connection.db;
  
  // Initialize GridFSBucket for Word documents (fs)
  wordGridfsBucket = new GridFSBucket(db, { bucketName: 'fs' }); // For Word documents

  // Initialize GridFSBucket for Excel files (excel_files)
  excelGridfsBucket = new GridFSBucket(db, { bucketName: 'excel_files' }); // For Excel files

  // Initialize GridFSBucket for Standard Contents (std_contents)
  stdContentsBucket = new GridFSBucket(db, { bucketName: 'std_contents' }); // For Word template
  
  console.log('GridFSBuckets initialized.');
});

// Endpoint to upload generated document
app.post('/api/upload-generated-document', upload.single('file'), (req, res) => {
  if (!req.file) {
    console.error('No file uploaded.');
    return res.status(400).send('No file uploaded.');
  }

  const readableStream = new Readable();
  readableStream.push(req.file.buffer);
  readableStream.push(null);

  const filename = req.file.originalname;
  const writeStream = wordGridfsBucket.openUploadStream(filename, {
    contentType: req.file.mimetype,
  });

  readableStream.pipe(writeStream);

  // Listen to the 'finish' event to get the file information
  writeStream.on('finish', () => {
    console.log(`File uploaded successfully, ID: ${writeStream.id}`);
    res.status(200).send({ fileId: writeStream.id });
  });

  writeStream.on('error', (err) => {
    console.error('Error uploading file:', err);
    res.status(500).send(err);
  });
});




// Endpoint to get the Word template using async/await
app.get('/api/template', async (req, res) => {
  const filename = 'OP###_Client_GridOS-DF_SER_v0.0.docx';
  console.log('Fetching template with filename:', filename);

  try {
    const files = await stdContentsBucket.find({ filename }).toArray();
    
    if (!files || files.length === 0) {
      console.error('Template not found');
      return res.status(404).json({ message: 'Template not found' });
    }

    const file = files[0];
    console.log('Template found:', file);

    const downloadStream = stdContentsBucket.openDownloadStream(file._id);

    res.set({
      'Content-Type': file.contentType,
      'Content-Disposition': `attachment; filename="${file.filename}"`,
    });

    downloadStream.on('error', (err) => {
      console.error('Error reading file from GridFS:', err);
      res.status(500).json({ message: 'Error reading file' });
    });

    downloadStream.on('finish', () => {
      console.log('File successfully sent to the client.');
    });

    downloadStream.pipe(res);
  } catch (err) {
    console.error('Error in /api/template endpoint:', err);
    res.status(500).json({ message: 'Server error' });
  }
});




// User Authentication Endpoint
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body;
  try {
    const user = await User.findOne({ username });
    if (!user) {
      console.log('User not found:', username);
      return res.status(400).json({ message: 'User not found' });
    }
    const isPasswordValid = await bcrypt.compare(password, user.password);
    if (!isPasswordValid) {
      console.log('Invalid password for user:', username);
      return res.status(400).json({ message: 'Invalid password' });
    }

    // Create a new log entry
    const logEntry = {
      username: user.username,
      loginTime: new Date(), // Current timestamp
      action: 'login'
    };

    // Insert the log entry into the 'User_logs' collection
    await UserLogs.create(logEntry);
    console.log('Login successful for user:', username);

    res.json({ message: 'Login successful', user });
  } catch (error) {
    console.error('Error during login:', error);
    res.status(500).json({ message: 'Server error' });
  }
});

// Endpoint to log user logout
app.post('/api/logout', async (req, res) => {
  const { username } = req.body;
  try {
    // Log user logout
    await UserLogs.create({
      username: username,
      loginTime: new Date(),
      action: 'logout'
    });
    console.log('Logout successful for user:', username);

    res.json({ message: 'Logout successful' });
  } catch (error) {
    console.error('Error during logout:', error);
    res.status(500).json({ message: 'Server error' });
  }
});

// Endpoint to get online users count and their usernames
app.get('/api/active-users', async (req, res) => {
  const now = new Date();
  const fiveMinutesAgo = new Date(now.getTime() - TIMEOUT_DURATION);

  try {
    // Fetch users who logged in within the last 5 minutes and have not logged out
    const onlineUsers = await UserLogs.aggregate([
      {
        $match: {
          loginTime: { $gte: fiveMinutesAgo },
          action: 'login'
        }
      },
      {
        $lookup: {
          from: 'userlogs',
          let: { username: '$username' },
          pipeline: [
            { $match: { $expr: { $and: [{ $eq: ['$username', '$$username'] }, { $eq: ['$action', 'logout'] }, { $gte: ['$loginTime', fiveMinutesAgo] }] } } }
          ],
          as: 'logoutLogs'
        }
      },
      {
        $match: {
          'logoutLogs.0': { $exists: false }
        }
      },
      {
        $group: {
          _id: "$username"
        }
      }
    ]).exec();

    const onlineUsernames = onlineUsers.map(user => user._id);
    const onlineUsersCount = onlineUsernames.length;

    console.log('Online users count:', onlineUsersCount);
    console.log('Online users:', onlineUsernames);

    res.json({ count: onlineUsersCount, users: onlineUsernames });
  } catch (error) {
    console.error('Error fetching online users:', error);
    res.status(500).json({ message: 'Server error' });
  }
});

app.get('/api/document/:id', async (req, res) => {
  const fileId = req.params.id;  // Document ID from frontend
  try {
    const file = await mongoose.connection.db.collection('fs.files')
      .findOne({ _id: mongoose.Types.ObjectId(fileId) });

    if (!file) {
      return res.status(404).json({ message: 'File not found' });
    }

    const downloadStream = gridfsBucket.openDownloadStream(file._id);

    res.setHeader('Content-Disposition', `inline; filename="${file.filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    
    downloadStream.pipe(res);
  } catch (error) {
    console.error('Error fetching file:', error);
    res.status(500).json({ message: 'Server error' });
  }
});






// Data Saving Endpoint
app.post('/api/data', async (req, res) => {
  const data = new Data(req.body);
  try {
    const savedData = await data.save();
    res.json({ message: 'Data saved successfully', savedData });
  } catch (error) {
    console.error('Error saving data:', error);
    res.status(500).json({ message: 'Server error' });
  }
});

// User creation endpoint
app.post('/api/register', async (req, res) => {
  const { username, password } = req.body;
  try {
    // Check if the user already exists
    const existingUser = await User.findOne({ username });
    if (existingUser) {
      return res.status(400).json({ message: 'User already exists' });
    }

    // Hash the password
    const hashedPassword = await bcrypt.hash(password, 10);

    // Create a new user
    const newUser = new User({ username, password: hashedPassword });
    await newUser.save();

    res.status(201).json({ message: 'User created successfully' });
  } catch (error) {
    console.error('Error creating user:', error);
    res.status(500).json({ message: 'Server error' });
  }
});

// Fetch all users
app.get('/api/users', async (req, res) => {
  try {
    const users = await User.find();
    res.json(users);
  } catch (error) {
    console.error('Failed to fetch users:', error);
    res.status(500).json({ message: 'Failed to fetch users' });
  }
});

// Delete users
app.post('/api/delete', async (req, res) => {
  const { userIds } = req.body;
  try {
    const validObjectIds = userIds.filter(id => mongoose.Types.ObjectId.isValid(id));
    if (validObjectIds.length !== userIds.length) {
      return res.status(400).json({ message: 'Invalid user IDs provided' });
    }
    const objectIdArray = validObjectIds.map(id => new mongoose.Types.ObjectId(id));
    await User.deleteMany({ _id: { $in: objectIdArray } });
    res.json({ message: 'Users deleted successfully' });
  } catch (error) {
    console.error('Error deleting users:', error);
    res.status(500).json({ message: error.message || 'Error deleting users' });
  }
});

// Trigger download endpoint
app.get('/trigger-download', async (req, res) => {
  try {
    const filenameToFind = 'processed_document_with_link.docx'; // Modify if needed
    const file = await mongoose.connection.db.collection('fs.files')
      .find({ filename: filenameToFind })
      .sort({ uploadDate: -1 })
      .limit(1)
      .next();

    if (!file) {
      console.log('No file found');
      return res.status(404).json({ message: 'No file found' });
    }

    console.log(`Found file: ${file.filename}`);

    const downloadStream = gridfsBucket.openDownloadStream(file._id);

    res.setHeader('Content-Disposition', `attachment; filename="${file.filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    downloadStream.pipe(res);

    downloadStream.on('error', (err) => {
      console.error('Error downloading file:', err);
      res.status(500).json({ message: 'Error downloading file' });
    });

    downloadStream.on('end', () => {
      res.end();
    });

  } catch (error) {
    console.error('Error triggering download:', error);
    res.status(500).json({ message: 'Server error' });
  }
});

// Fetch logs
app.get('/api/logs', async (req, res) => {
  try {
    const logs = await mongoose.connection.db.collection('userlogs')
      .find({})
      .sort({ startTime: -1 })
      .toArray();

    console.log('Logs data:', logs);

    res.setHeader('Content-Type', 'application/json');
    res.json(logs);
  } catch (error) {
    console.error('Error fetching logs:', error);
    res.status(500).json({ message: 'Server error fetching logs' });
  }
});

// Endpoint to log document download
app.post('/api/document-processed', async (req, res) => {
  const { username, timestamp } = req.body;

  console.log('Received data:', { username, timestamp });

  try {
    // Access the MongoDB collection directly
    const collection = mongoose.connection.db.collection('documentprocessed');

    // Insert a new document into the 'documentprocesseds' collection
    const result = await collection.insertOne({
      username,
      timestamp
    });

    // Check if the insertion was successful
    if (result.acknowledged) {
      res.status(201).json({ message: 'Document download logged successfully', insertedId: result.insertedId });
    } else {
      res.status(500).json({ message: 'Failed to log document download' });
    }
  } catch (error) {
    console.error('Error logging document download:', error);
    res.status(500).json({ message: 'Server error logging document download' });
  }
});

// Endpoint to log button click events
app.post('/api/button-clicked', async (req, res) => {
  const { username, timestamp, action } = req.body;

  console.log('Button click logged:', { username, timestamp, action });

  try {
    // Access the MongoDB collection directly
    const collection = mongoose.connection.db.collection('buttonclicklogs');

    // Insert a new document into the 'buttonclicklogs' collection
    const result = await collection.insertOne({
      username,
      timestamp,
      action
    });

    if (result.acknowledged) {
      res.status(201).json({ message: 'Button click logged successfully', insertedId: result.insertedId });
    } else {
      res.status(500).json({ message: 'Failed to log button click' });
    }
  } catch (error) {
    console.error('Error logging button click:', error);
    res.status(500).json({ message: 'Server error logging button click' });
  }
});



// Fetch button click logs
app.get('/api/button-click-logs', async (req, res) => {
  try {
    const logs = await mongoose.connection.db.collection('buttonclicklogs')
      .find({})
      .sort({ timestamp: -1 }) // Sort by timestamp, descending
      .toArray();

    res.json(logs);
  } catch (error) {
    console.error('Error fetching button click logs:', error);
    res.status(500).json({ message: 'Error fetching button click logs' });
  }
});






// Initialize GridFSBucket for Word documents (fs)
mongoose.connection.once('open', () => {
  const db = mongoose.connection.db;
  wordGridfsBucket = new GridFSBucket(db, { bucketName: 'fs' }); // For Word documents
});

// Initialize GridFSBucket for Excel files (excel_files)
mongoose.connection.once('open', () => {
  const db = mongoose.connection.db;
  excelGridfsBucket = new GridFSBucket(db, { bucketName: 'excel_files' }); // For Excel files
});


// Fetch the latest Excel file from GridFS
app.get('/api/excel-file', async (req, res) => {
  try {
    const file = await mongoose.connection.db.collection('excel_files.files')
      .find({})
      .sort({ uploadDate: -1 })
      .limit(1)
      .next();

    if (!file) {
      return res.status(404).json({ message: 'No Excel file found' });
    }

    console.log(`Found Excel file: ${file.filename}`);

    const downloadStream = gridfsBucket.openDownloadStream(file._id);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${file.filename}"`);

    downloadStream.pipe(res);

    downloadStream.on('error', (err) => {
      if (err.code === 'ENOENT') {
        console.error('File not found in chunks collection:', err);
        return res.status(404).json({ message: 'File chunks not found, possibly corrupted or not fully uploaded' });
      }
      console.error('Error downloading file:', err);
      res.status(500).json({ message: 'Error downloading file' });
    });

    downloadStream.on('end', () => {
      res.end();
    });
  } catch (error) {
    console.error('Error fetching Excel file:', error);
    res.status(500).json({ message: 'Server error fetching Excel file' });
  }
});

// Route to generate document URL and JWT token
app.post('/api/generate-editor-url1', (req, res) => {
  const documentKey = generateDocumentKey();
  const documentTitle = `Document_${Date.now()}`;

  const document = {
    fileType: "docx",
    key: documentKey,
    title: documentTitle,
    url: "http://192.168.50.57:8081/document", // Updated URL
    permissions: { edit: true, review: true }
  };

  const editorConfig = {
    mode: "edit",
    user: {
      id: "user123",
      name: "User"
    }
  };

  const config = {
    document,
    documentType: "word",
    editorConfig
  };

  // Generate the JWT token with the entire config as payload
  const token = jwt.sign(config, jwtSecretKey, { expiresIn: '1h' });

  console.log('Generated JWT Token:', token);

  res.json({ document, token });
});

// Route to generate editor URL and JWT token
app.post('/api/generate-editor-url', async (req, res) => {
  const { fileId } = req.body;

  if (!fileId) {
    return res.status(400).json({ error: 'fileId is required' });
  }

  try {
    // Correct the ObjectId creation
    const _id = new mongoose.Types.ObjectId(fileId);

    // Fetch the file from GridFS
    const files = await wordGridfsBucket.find({ _id }).toArray();

    if (files.length === 0) {
      return res.status(404).json({ error: 'File not found' });
    }

    const file = files[0];

    // Generate a unique document key instead of using fileId directly
    const documentKey = generateDocumentKey(); // Generate the document key like in the working setup
    
    // Update the document URL to use port 8081
    const documentUrl = `http://192.168.50.57:8081/document/${fileId}`;

    const document = {
      fileType: "docx",
      key: documentKey, // Use generated key here
      title: file.filename,
      url: documentUrl,
      permissions: { edit: true, review: true },
    };

    const editorConfig = {
      mode: "edit",
      user: {
        id: "user123",
        name: "User",
      },
    };

    const config = {
      document,
      documentType: "word",
      editorConfig,
    };

    // Generate the JWT token with the entire config as payload
  const token = jwt.sign(config, jwtSecretKey, { expiresIn: '1h' });

    console.log('Generated Dynamic JWT Token:', token);
    res.json({ document, token });
  } catch (error) {
    console.error('Error generating editor URL:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});



// Route to serve the document file
app.get('/document/:fileId', (req, res) => {
  const authHeader = req.headers['authorization'];

  if (!authHeader) {
    return res.status(401).json({ error: 'No token provided' });
  }

  const token = authHeader.replace('Bearer ', '');

  // Verify the JWT token
  jwt.verify(token, jwtSecretKey, (err, decoded) => {
    if (err) {
      console.error('JWT Verification Error:', err);
      return res.status(401).json({ error: 'Invalid token' });
    }

    const { fileId } = req.params;

    // OPTIONAL: If document key validation is necessary, ensure it matches
    if (decoded.document && decoded.document.key !== fileId) {
      return res.status(403).json({ error: 'Access denied: document key mismatch' });
    }

    // Correct ObjectId creation using 'new'
    const downloadStream = wordGridfsBucket.openDownloadStream(new mongoose.Types.ObjectId(fileId));

    downloadStream.on('error', (err) => {
      console.error('Error downloading file:', err);
      res.status(404).json({ error: 'File not found' });
    });

    // Set the appropriate Content-Type header for a DOCX file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="downloaded-file.docx"`);

    // Pipe the file to the response
    downloadStream.pipe(res);
  });
});




// Route to serve the document file
app.get('/document', (req, res) => {
  const authHeader = req.headers['authorization'];

  if (!authHeader) {
    return res.status(401).json({ error: 'No token provided' });
  }

  const token = authHeader.replace('Bearer ', '');

  // Verify the JWT token
  jwt.verify(token, jwtSecretKey, (err, decoded) => {
    if (err) {
      console.error('JWT Verification Error:', err);
      return res.status(401).json({ error: 'Invalid token' });
    }

    // Path to your document file
    const filePath = 'C:\\Users\\rikog\\Downloads\\OP###_Client_GridOS-DF_SER_v0.0.docx';
    

    // Check if the file exists
    if (!fs.existsSync(filePath)) {
      console.error('File not found:', filePath);
      return res.status(404).json({ error: 'File not found' });
    }

    // Set the appropriate Content-Type header for a DOCX file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="OP###_Client_GridOS-DF_SER_v0.0.docx"');

    // Send the file
    res.sendFile(filePath, (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).send('Error sending file');
      } else {
        console.log('File sent successfully:', filePath);
      }
    });
  });
});

app.get('/api/get-excel', (req, res) => {
  const authHeader = req.headers['authorization'];

  if (!authHeader) {
    return res.status(401).json({ error: 'No token provided' });
  }

  const token = authHeader.replace('Bearer ', '');

  // Verify the JWT token
  jwt.verify(token, jwtSecretKey, async (err, decoded) => {
    if (err) {
      console.error('JWT Verification Error:', err);
      return res.status(401).json({ error: 'Invalid token' });
    }

    try {
      const filename = 'RACI_Table.xlsx'; // Hardcoded filename

      const file = await mongoose.connection.db.collection('excel_files.files')
        .findOne({ filename });

      if (!file) {
        return res.status(404).json({ message: 'Excel file not found' });
      }

      // Use excelGridfsBucket to open the download stream
      const downloadStream = excelGridfsBucket.openDownloadStream(file._id);

      res.setHeader('Content-Disposition', `inline; filename="${file.filename}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

      downloadStream.pipe(res);

      downloadStream.on('error', (err) => {
        console.error('Error downloading Excel file:', err);
        res.status(500).json({ message: 'Error streaming Excel document' });
      });

      downloadStream.on('end', () => {
        res.end();
      });
    } catch (error) {
      console.error('Error fetching Excel document:', error);
      res.status(500).json({ message: 'Server error' });
    }
  });
});


// Route to generate Excel document URL and JWT token
app.post('/api/generate-excel-editor-url1', (req, res) => {
  const documentKey = generateDocumentKey();
  const documentTitle = 'RACI_Table.xlsx';

  const document = {
    fileType: "xlsx",
    key: documentKey,
    title: documentTitle,
    url: "http://192.168.50.57:8081/api/get-excel",
    permissions: { edit: true, review: true }
  };

  const editorConfig = {
    mode: "edit",
    user: {
      id: "user123",
      name: "User"
    }
  };

  const config = {
    document,
    documentType: "cell",
    editorConfig
  };

  // Generate the JWT token with the entire config as payload
  const token = jwt.sign(config, jwtSecretKey, { expiresIn: '1h' });

  console.log('Generated JWT Token for Excel:', token);

  // Return { document, token } instead of { config, token }
  res.json({ document, token });
});

// Route to generate editor URL and JWT token for Excel files
app.post('/api/generate-editor-url-excel', async (req, res) => {
  const { fileId } = req.body;

  if (!fileId) {
    return res.status(400).json({ error: 'fileId is required' });
  }

  try {
    // Fetch the file from GridFS
    const _id = new mongoose.Types.ObjectId(fileId);
    const files = await excelGridfsBucket.find({ _id }).toArray();

    if (files.length === 0) {
      return res.status(404).json({ error: 'File not found' });
    }

    const file = files[0];
    
    // Generate a unique document key for the Excel file
    const documentKey = generateDocumentKey();
    
    // Set the URL to access the Excel file in ONLYOFFICE
    const documentUrl = `http://192.168.50.57:8081/excel/${fileId}`;

    const document = {
      fileType: "xlsx",  // Excel file type
      key: documentKey,
      title: file.filename,
      url: documentUrl,
      permissions: { edit: true, review: true },
    };

    const editorConfig = {
      mode: "edit",
      user: {
        id: "user123",
        name: "User",
      },
    };

    const config = {
      document,
      documentType: "cell",  // Specify Excel (cell) document type
      editorConfig,
    };

    // Generate the JWT token for Excel files
    const token = jwt.sign(config, jwtSecretKey, { expiresIn: '1h' });

    res.json({ document, token });
  } catch (error) {
    console.error('Error generating editor URL for Excel:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});


// Route to serve the Excel file
app.get('/excel/:fileId', (req, res) => {
  const authHeader = req.headers['authorization'];

  if (!authHeader) {
    return res.status(401).json({ error: 'No token provided' });
  }

  const token = authHeader.replace('Bearer ', '');

  // Verify the JWT token
  jwt.verify(token, jwtSecretKey, (err, decoded) => {
    if (err) {
      console.error('JWT Verification Error:', err);
      return res.status(401).json({ error: 'Invalid token' });
    }

    const { fileId } = req.params;

    // **Change this check**: If document key validation is necessary, ensure it matches the fileId
    // Instead of comparing the fileId, compare it with the document's actual fileId, not documentKey
    if (decoded.document && decoded.document.url.split('/').pop() !== fileId) {
      return res.status(403).json({ error: 'Access denied: document key mismatch' });
    }

    // Fetch the Excel file from MongoDB
    const downloadStream = excelGridfsBucket.openDownloadStream(new mongoose.Types.ObjectId(fileId));

    downloadStream.on('error', (err) => {
      console.error('Error downloading Excel file:', err);
      res.status(404).json({ error: 'File not found' });
    });

    // Set the appropriate Content-Type header for Excel files
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="downloaded-excel.xlsx"`);

    // Pipe the Excel file to the response
    downloadStream.pipe(res);
  });
});







// Start the server
const PORT = 8081;
app.listen(PORT, () => {
  console.log(`Server is running on http://192.168.50.57:${PORT}`);
});
