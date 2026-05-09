import express from 'express';
import cors from 'cors';
import metricsRouter from './routes/metricsRoutes.js';
import { errorHandler, notFoundHandler } from './middleware/errorHandler.js';

const app = express();
const allowedOrigins = (
  process.env.ALLOWED_ORIGINS ||
  [
    'http://localhost:3000',
    'http://localhost:5173',
    'http://127.0.0.1:3000',
    'http://127.0.0.1:5173',
    'https://acadsimulationfrontend.vercel.app',
  ].join(',')
)
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

app.use(
  cors({
    origin(origin, callback) {
      if (!origin || allowedOrigins.includes(origin)) {
        return callback(null, true);
      }

      return callback(new Error(`Origin ${origin} is not allowed by CORS.`));
    },
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type'],
  }),
);
app.use(express.json());

// app.get("/", (_req, res) => {
//   res.json({
//     status: "ok",
//     service: "Academic Block Simulation Backend",
//     endpoints: ["/health", "/metrics", "/history", "/upload", "/simulate", "/optimize"]
//   });
// });

app.get('/health', (_req, res) => {
  res.json({ status: 'ok' });
});

app.use('/', metricsRouter);
app.use(notFoundHandler);
app.use(errorHandler);

export default app;
