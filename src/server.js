import app from './app.js';
import { initializePersistence, isMongoConfigured } from "./store/persistence.js";

const PORT = Number(process.env.PORT) || 8000;
const HOST = "0.0.0.0";

await initializePersistence();

app.listen(PORT, HOST, () => {
  console.log(`Academic Simulation backend listening on ${HOST}:${PORT}`);
  console.log(
    `Persistence backend: ${isMongoConfigured() ? "MongoDB Atlas" : "local filesystem"}`
  );
});

const axios = require('axios');

const URL = 'https://acadsimulationbackend-80sa.onrender.com/';

setInterval(async () => {
  try {
    const response = await axios.get(URL);
    console.log('Pinged successfully:', response.status);
  } catch (err) {
    console.log('Ping failed');
  }
}, 30000);
