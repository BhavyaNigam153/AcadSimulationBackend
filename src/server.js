import app from './app.js';
import axios from "axios";
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



