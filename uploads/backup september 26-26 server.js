// server.js
const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const bcrypt = require('bcrypt');
const { MongoClient, GridFSBucket } = require('mongodb');
const app = express();
const fs = require('fs').promises; // Using fs.promises for async/await

// Middleware
app.use(bodyParser.json());
app.use(cors({
  origin: 'http://192.168.50.57:8082',
  methods: 'GET,POST,PUT,DELETE',
  credentials: true,
}));

app.use('/onlyoffice-api', express.static('path/to/your/onlyoffice-api.js', {
  setHeaders: (res, path) => {
    res.setHeader('Content-Type', 'application/javascript');
  }
}));

// MongoDB connection
const mongoURI = 'mongodb://192.168.50.57:27017/local';
mongoose.connect(mongoURI, { useNewUrlParser: true, useUnifiedTopology: true });

mongoose.connection.on('connected', () => {
  console.log('Connected to MongoDB');
});

mongoose.connection.on('error', (err) => {
  console.log('Error connecting to MongoDB:', err);
});

// Initialize GridFSBucket only after the connection is successful
let gridfsBucket;
mongoose.connection.once('open', () => {
  const db = mongoose.connection.db;
  gridfsBucket = new GridFSBucket(db, { bucketName: 'fs' });
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

// Initialize GridFSBucket only after the connection is successful

mongoose.connection.once('open', () => {
  const db = mongoose.connection.db;
  gridfsBucket = new GridFSBucket(db, { bucketName: 'excel_files' });
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





// Start the server
const PORT = 8081;
app.listen(PORT, () => {
  console.log(`Server is running on http://192.168.50.57:${PORT}`);
});
