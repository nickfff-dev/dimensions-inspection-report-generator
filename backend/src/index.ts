import express, { Express } from 'express';
import cors from 'cors';
import path from 'path';

// Import routes and middleware
import apiRoutes from './routes/api';
import { errorHandler } from './middleware/validation';

const app: Express = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Static files for serving generated Excel files
app.use('/generated', express.static(path.join(__dirname, '../generated')));

// Routes
app.use('/api', apiRoutes);

// Health check
app.get('/health', (_, res) => {
  res.json({ status: 'OK', message: 'Excel Generation API is running' });
});

// Error handling middleware (must be last)
app.use(errorHandler);

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Backend server running on http://localhost:${PORT}`);
  console.log(`📁 Generated files will be served from /generated`);
});

export default app;