import { MongoClient } from "mongodb";

let mongoClientPromise = null;

export function isMongoConfigured() {
  return Boolean(process.env.MONGODB_URI?.trim());
}

export async function initializePersistence() {
  if (!isMongoConfigured()) {
    return null;
  }

  const collection = await getRecordsCollection();
  await collection.createIndex({ id: 1 }, { unique: true });
  await collection.createIndex({ dataDate: -1, uploadedAt: -1 });
  return collection;
}

export async function getRecordsCollection() {
  if (!isMongoConfigured()) {
    return null;
  }

  if (!mongoClientPromise) {
    const mongoUri = process.env.MONGODB_URI.trim();
    const dbName = process.env.MONGODB_DB_NAME?.trim() || "acad_simulation";
    const client = new MongoClient(mongoUri);

    mongoClientPromise = client.connect().then((connectedClient) => ({
      client: connectedClient,
      db: connectedClient.db(dbName)
    }));
  }

  const { db } = await mongoClientPromise;
  return db.collection("records");
}
