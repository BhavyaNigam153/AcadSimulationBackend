import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  buildDatasetFromWorkbook,
  buildMetricsFromDataset,
} from "../services/excelService.js";
import { getRecordsCollection, isMongoConfigured } from "./persistence.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const dataDirectory = path.resolve(__dirname, "../../data");
const uploadsDirectory = path.join(dataDirectory, "uploads");
const recordsDirectory = path.join(dataDirectory, "records");
const baseTemplatePath = path.resolve(__dirname, "../../../23_03_2026.xlsx");

export async function saveStoredWorkbook({ originalName, buffer, dataset, metrics }) {
  await ensureStorage();

  const now = new Date();
  const dataDate = parseDataDateFromFileName(originalName || "");
  const id = `record-${now.getTime()}`;
  const safeName = sanitizeFileName(originalName || "upload.xlsx");
  const workbookFileName = `${id}-${safeName}`;
  const storedFileName = isMongoConfigured() ? null : workbookFileName;

  if (!isMongoConfigured()) {
    const workbookPath = path.join(uploadsDirectory, workbookFileName);
    await fs.writeFile(workbookPath, buffer);
  }

  const normalizedDataset = normalizeDataset(
    dataset || (await buildDatasetFromWorkbook(buffer)),
  );
  const resolvedMetrics =
    metrics || buildMetricsFromDataset(normalizedDataset);

  const record = buildStoredRecord({
    id,
    originalName: originalName || "upload.xlsx",
    storedFileName,
    uploadedAt: now.toISOString(),
    dataDate,
    metrics: resolvedMetrics,
    dataset: normalizedDataset,
    source: "upload",
  });

  await writeRecord(record);

  return record.metrics;
}

export async function saveTimetableRecord({
  dataDate,
  timetable,
  basedOnRecordId = null,
  targetRecordId = null,
}) {
  await ensureStorage();

  const baseDataset = basedOnRecordId
    ? await getDatasetForRecord(basedOnRecordId)
    : await getBaseTemplateDataset();
  const normalizedDataDate = normalizeInputDate(dataDate);
  const existingRecord =
    (targetRecordId ? await getRecordById(targetRecordId) : null) ||
    (await findRecordByDate(normalizedDataDate));
  const dataset = normalizeDataset({
    rooms: baseDataset.rooms,
    enrollments: baseDataset.enrollments,
    timetable,
  });
  const metrics = buildMetricsFromDataset(dataset);
  const now = new Date();
  const id = existingRecord?.id || `record-${now.getTime()}`;
  const originalName =
    existingRecord?.originalName || formatWorkbookNameFromIsoDate(normalizedDataDate);

  const record = buildStoredRecord({
    id,
    originalName,
    storedFileName: existingRecord?.storedFileName || null,
    uploadedAt: now.toISOString(),
    dataDate: normalizedDataDate,
    metrics,
    dataset,
    source: existingRecord?.source || "timetable-editor",
  });

  await writeRecord(record);

  return record.metrics;
}

export async function getStoredMetrics() {
  const records = await readAllRecords();
  return records[0]?.metrics || null;
}

export async function listStoredHistory() {
  const records = await readAllRecords();
  return records.map((record) => ({
    id: record.id,
    originalName: record.originalName,
    uploadedAt: record.uploadedAt,
    dataDate: normalizeRecordDate(record),
    source: record.source || "upload",
    weekday: getWeekday(normalizeRecordDate(record)),
    summary: record.metrics.summary,
  }));
}

export async function getStoredRecord(recordId) {
  const records = await readAllRecords();
  const record = records.find((entry) => entry.id === recordId);
  return record?.metrics || null;
}

export async function getTimetableTemplate(recordId = null) {
  const baseDataset = recordId
    ? await getDatasetForRecord(recordId)
    : await getBaseTemplateDataset();

  return {
    sourceRecordId: recordId,
    templateDate: recordId ? null : "2026-03-23",
    rooms: baseDataset.rooms.map((room) => ({
      roomId: room.roomId,
      roomNameEn: room.roomNameEn,
      roomNameHi: room.roomNameHi,
      floor: room.floor,
      type: room.type,
      capacity: room.capacity,
    })),
    timetable: baseDataset.timetable.map((entry) => ({
      classId: entry.classId,
      subject: entry.subject,
      roomId: entry.roomId,
      startTime: entry.startTime,
      endTime: entry.endTime,
      day: entry.day,
    })),
  };
}

export async function applyOptimizationToRecord({
  recordId,
  classId,
  day = null,
  startTime = null,
  toRoom,
}) {
  const sourceRecord = recordId
    ? await getRecordById(recordId)
    : (await readAllRecords())[0] || null;

  if (!sourceRecord) {
    const error = new Error("No source timetable record found to apply optimization.");
    error.statusCode = 404;
    throw error;
  }

  const dataset = await hydrateDatasetFromRecord(sourceRecord);
  const roomExists = dataset.rooms.some((room) => room.roomId === toRoom);
  if (!roomExists) {
    const error = new Error(`Target room ${toRoom} does not exist in the stored room inventory.`);
    error.statusCode = 400;
    throw error;
  }

  const entryIndex = dataset.timetable.findIndex(
    (entry) =>
      entry.classId === classId &&
      (day ? entry.day === day : true) &&
      (startTime ? entry.startTime === startTime : true),
  );
  if (entryIndex < 0) {
    const error = new Error(`Class ${classId} was not found in the stored timetable.`);
    error.statusCode = 404;
    throw error;
  }

  const selectedEntry = dataset.timetable[entryIndex];
  const hasConflict = dataset.timetable.some((entry, index) => {
    if (index === entryIndex) {
      return false;
    }

    if (entry.roomId !== toRoom || entry.day !== selectedEntry.day) {
      return false;
    }

    return (
      toMinutes(selectedEntry.startTime) < toMinutes(entry.endTime) &&
      toMinutes(selectedEntry.endTime) > toMinutes(entry.startTime)
    );
  });

  if (hasConflict) {
    const error = new Error(
      `Room ${toRoom} already has a class assigned during ${selectedEntry.day} ${selectedEntry.startTime}-${selectedEntry.endTime}.`,
    );
    error.statusCode = 400;
    throw error;
  }

  const updatedTimetable = dataset.timetable.map((entry, index) =>
    index === entryIndex ? { ...entry, roomId: toRoom } : entry,
  );

  return saveTimetableRecord({
    dataDate: normalizeRecordDate(sourceRecord),
    timetable: updatedTimetable,
    basedOnRecordId: sourceRecord.id,
    targetRecordId: sourceRecord.id,
  });
}

export async function buildPrediction(daysAhead) {
  const records = await readAllRecords();
  if (records.length === 0) {
    return null;
  }

  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + daysAhead);
  const targetDateIso = toDateOnlyIso(targetDate);
  const targetWeekday = getWeekday(targetDateIso);

  const sameWeekdayRecords = records.filter(
    (record) => getWeekday(normalizeRecordDate(record)) === targetWeekday,
  );
  const fallbackRecords = records.slice(0, Math.min(records.length, 5));
  const sourceRecords = sameWeekdayRecords.length > 0 ? sameWeekdayRecords : fallbackRecords;
  const predictionStrategy =
    sameWeekdayRecords.length > 0 ? "same-weekday-history" : "recent-history-fallback";

  return attachPredictionMeta(computePredictedMetrics(sourceRecords), {
    targetDate: targetDateIso,
    targetWeekday,
    daysAhead,
    predictionStrategy,
    basedOnRecordCount: sourceRecords.length,
    basedOnRecordIds: sourceRecords.map((record) => record.id),
  });
}

async function ensureDirectories() {
  await fs.mkdir(uploadsDirectory, { recursive: true });
  await fs.mkdir(recordsDirectory, { recursive: true });
}

async function ensureStorage() {
  if (isMongoConfigured()) {
    await getRecordsCollection();
    return;
  }

  await ensureDirectories();
}

async function writeRecord(record) {
  if (isMongoConfigured()) {
    const collection = await getRecordsCollection();
    await collection.replaceOne({ id: record.id }, sanitizeRecordForMongo(record), {
      upsert: true
    });
    return;
  }

  await fs.writeFile(path.join(recordsDirectory, `${record.id}.json`), JSON.stringify(record, null, 2), "utf8");
}

async function readAllRecords() {
  await ensureStorage();

  const records = isMongoConfigured()
    ? await readMongoRecords()
    : await readSavedUploadRecords();

  return records.sort((left, right) => {
    const dateDiff =
      new Date(normalizeRecordDate(right)).getTime() -
      new Date(normalizeRecordDate(left)).getTime();
    if (dateDiff !== 0) {
      return dateDiff;
    }

    return new Date(right.uploadedAt).getTime() - new Date(left.uploadedAt).getTime();
  });
}

async function readSavedUploadRecords() {
  const fileNames = await fs.readdir(recordsDirectory);
  const jsonFiles = fileNames.filter((fileName) => fileName.endsWith(".json"));

  return Promise.all(
    jsonFiles.map(async (fileName) => {
      const contents = await fs.readFile(path.join(recordsDirectory, fileName), "utf8");
      return JSON.parse(contents);
    }),
  );
}

async function readMongoRecords() {
  const collection = await getRecordsCollection();
  const records = await collection.find({}, { projection: { _id: 0 } }).toArray();
  return records.map(stripMongoId);
}

async function getRecordById(recordId) {
  const records = await readAllRecords();
  return records.find((record) => record.id === recordId) || null;
}

async function findRecordByDate(dataDate) {
  const records = await readAllRecords();
  return records.find((record) => normalizeRecordDate(record) === dataDate) || null;
}

async function getDatasetForRecord(recordId) {
  const record = await getRecordById(recordId);
  if (!record) {
    const error = new Error("Requested template record was not found.");
    error.statusCode = 404;
    throw error;
  }

  return hydrateDatasetFromRecord(record);
}

async function getBaseTemplateDataset() {
  try {
    const workbookBuffer = await fs.readFile(baseTemplatePath);
    return normalizeDataset(await buildDatasetFromWorkbook(workbookBuffer));
  } catch (_error) {
    const latestRecord = (await readAllRecords())[0] || null;
    if (latestRecord) {
      return hydrateDatasetFromRecord(latestRecord);
    }

    const error = new Error("No base template workbook is available for timetable editing.");
    error.statusCode = 404;
    throw error;
  }
}

async function hydrateDatasetFromRecord(record) {
  if (record.dataset) {
    return normalizeDataset(record.dataset);
  }

  const metrics = record.metrics || {};
  const rooms = (metrics.rooms || []).map((room) => ({
    roomId: room.roomId,
    roomNameEn: room.roomNameEn || "",
    roomNameHi: room.roomNameHi || "",
    capacity: room.capacity || 0,
    type: room.type || "",
    floor: room.floor || "",
  }));

  const timetable = (metrics.classes || []).map((classItem) => ({
    classId: classItem.classId,
    subject: classItem.subject,
    roomId: classItem.roomId,
    startTime: classItem.startTime,
    endTime: classItem.endTime,
    day: classItem.day,
  }));

  const enrollments = dedupeByKey(
    (metrics.classes || []).map((classItem) => ({
      classId: classItem.classId,
      studentCount: classItem.studentCount,
    })),
    "classId",
  );

  return normalizeDataset({ rooms, timetable, enrollments });
}

function buildStoredRecord({
  id,
  originalName,
  storedFileName,
  uploadedAt,
  dataDate,
  metrics,
  dataset,
  source,
}) {
  return {
    id,
    originalName,
    storedFileName,
    uploadedAt,
    dataDate,
    source,
    dataset,
    metrics: attachActualMeta(metrics, {
      id,
      originalName,
      uploadedAt,
      dataDate,
      source,
    }),
  };
}

function normalizeDataset(dataset) {
  return {
    rooms: (dataset.rooms || []).map((room) => ({
      roomId: String(room.roomId || "").trim(),
      roomNameEn: String(room.roomNameEn || "").trim(),
      roomNameHi: String(room.roomNameHi || "").trim(),
      capacity: Number(room.capacity) || 0,
      type: String(room.type || "").trim(),
      floor: String(room.floor || "").trim(),
    })),
    timetable: (dataset.timetable || []).map((entry) => ({
      classId: String(entry.classId || "").trim(),
      subject: String(entry.subject || "").trim(),
      roomId: String(entry.roomId || "").trim(),
      startTime: normalizeTime(entry.startTime),
      endTime: normalizeTime(entry.endTime),
      day: String(entry.day || "").trim(),
    })),
    enrollments: dedupeByKey(
      (dataset.enrollments || []).map((entry) => ({
        classId: String(entry.classId || "").trim(),
        studentCount: Number(entry.studentCount) || 0,
      })),
      "classId",
    ),
  };
}

function normalizeTime(value) {
  const normalized = String(value || "").trim();
  if (!normalized) {
    return "";
  }

  const [hours, minutes] = normalized.split(":").map(Number);
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function dedupeByKey(items, key) {
  const seen = new Map();
  for (const item of items) {
    seen.set(item[key], item);
  }
  return [...seen.values()];
}

function computePredictedMetrics(records) {
  const metricsList = records.map((record) => record.metrics);
  const roomMap = new Map();
  const timeSeriesMap = new Map();
  const latestClasses =
    records.find((record) => (record.metrics?.classes || []).length > 0)?.metrics?.classes || [];

  for (const metrics of metricsList) {
    for (const room of metrics.rooms || []) {
      const entry = roomMap.get(room.roomId) || {
        roomId: room.roomId,
        capacityTotal: 0,
        occupancyTotal: 0,
        utilizationTotal: 0,
        count: 0,
        recommendations: [],
        roomNamesEn: [],
        roomNamesHi: [],
        floors: [],
        types: [],
      };

      entry.capacityTotal += room.capacity || 0;
      entry.occupancyTotal += room.avgOccupancy || 0;
      entry.utilizationTotal += room.utilization || 0;
      entry.count += 1;
      if (room.roomNameEn) {
        entry.roomNamesEn.push(room.roomNameEn);
      }
      if (room.roomNameHi) {
        entry.roomNamesHi.push(room.roomNameHi);
      }
      if (room.floor) {
        entry.floors.push(room.floor);
      }
      if (room.type) {
        entry.types.push(room.type);
      }
      if (room.recommendation) {
        entry.recommendations.push(room.recommendation);
      }
      roomMap.set(room.roomId, entry);
    }

    for (const point of metrics.timeSeries || []) {
      const entry = timeSeriesMap.get(point.time) || {
        time: point.time,
        studentsTotal: 0,
        count: 0,
      };
      entry.studentsTotal += point.students || 0;
      entry.count += 1;
      timeSeriesMap.set(point.time, entry);
    }
  }

  const rooms = [...roomMap.values()]
    .map((entry) => {
      const capacity = Math.round(entry.capacityTotal / Math.max(entry.count, 1));
      const avgOccupancy = round(entry.occupancyTotal / Math.max(entry.count, 1));
      const utilization = round(entry.utilizationTotal / Math.max(entry.count, 1));

      let status = "normal";
      if (avgOccupancy > 1) {
        status = "overutilized";
      } else if (utilization < 0.4) {
        status = "underutilized";
      }

      return {
        roomId: entry.roomId,
        roomNameEn: mostCommon(entry.roomNamesEn),
        roomNameHi: mostCommon(entry.roomNamesHi),
        floor: mostCommon(entry.floors),
        type: mostCommon(entry.types),
        capacity,
        avgOccupancy,
        utilization,
        status,
        recommendation: mostCommon(entry.recommendations),
      };
    })
    .sort((left, right) => left.roomId.localeCompare(right.roomId));

  const timeSeries = [...timeSeriesMap.values()]
    .map((entry) => ({
      time: entry.time,
      students: Math.round(entry.studentsTotal / Math.max(entry.count, 1)),
    }))
    .sort((left, right) => toMinutes(left.time) - toMinutes(right.time));

  const avgOccupancy =
    metricsList.reduce((sum, metrics) => sum + (metrics.summary?.avgOccupancy || 0), 0) /
    Math.max(metricsList.length, 1);

  const totalRooms =
    metricsList.reduce((sum, metrics) => sum + (metrics.summary?.totalRooms || 0), 0) /
    Math.max(metricsList.length, 1);

  const peakHour = mostCommon(
    metricsList.map((metrics) => metrics.summary?.peakHour).filter(Boolean),
  );

  const classes = latestClasses.map((classItem) => {
    const room = rooms.find((roomEntry) => roomEntry.roomId === classItem.roomId);
    const capacity = room?.capacity || classItem.capacity || 0;
    const occupancy = room?.avgOccupancy || classItem.occupancy || 0;
    const studentCount = Math.round(capacity * occupancy);

    let status = "normal";
    if (occupancy > 1) {
      status = "overcrowded";
    } else if (occupancy < 0.4) {
      status = "underutilized";
    }

    return {
      ...classItem,
      roomNameEn: room?.roomNameEn || classItem.roomNameEn || "",
      roomNameHi: room?.roomNameHi || classItem.roomNameHi || "",
      floor: room?.floor || classItem.floor || "",
      type: room?.type || classItem.type || "",
      capacity,
      studentCount,
      occupancy: round(occupancy),
      status,
    };
  });

  return {
    summary: {
      totalRooms: Math.round(totalRooms),
      avgOccupancy: round(avgOccupancy),
      peakHour: peakHour || "N/A",
    },
    rooms,
    timeSeries,
    classes,
  };
}

function attachActualMeta(metrics, { id, originalName, uploadedAt, dataDate, source = "upload" }) {
  return {
    ...metrics,
    meta: {
      recordId: id,
      recordType: "actual",
      originalName,
      uploadedAt,
      dataDate,
      source,
      label: source === "timetable-editor" ? "Timetable Editor Save" : "Uploaded Workbook",
      weekday: getWeekday(dataDate),
    },
  };
}

function attachPredictionMeta(metrics, metadata) {
  return {
    ...metrics,
    meta: {
      recordType: "predicted",
      label: "Predicted Metrics",
      generatedAt: new Date().toISOString(),
      ...metadata,
    },
  };
}

function sanitizeFileName(fileName) {
  return fileName.replace(/[^a-zA-Z0-9._-]/g, "_");
}

function getWeekday(isoDate) {
  return new Date(`${isoDate}T00:00:00`).toLocaleDateString("en-US", { weekday: "long" });
}

function parseDataDateFromFileName(fileName) {
  const match = String(fileName).match(/^(\d{2})_(\d{2})_(\d{4})\.xlsx$/i);
  if (!match) {
    return null;
  }

  const [, day, month, year] = match;
  const isoDate = `${year}-${month}-${day}`;
  const parsedDate = new Date(`${isoDate}T00:00:00`);

  if (Number.isNaN(parsedDate.getTime())) {
    return null;
  }

  if (
    parsedDate.getFullYear() !== Number(year) ||
    parsedDate.getMonth() + 1 !== Number(month) ||
    parsedDate.getDate() !== Number(day)
  ) {
    return null;
  }

  return isoDate;
}

function normalizeInputDate(value) {
  const normalized = toDateOnlyIso(value);
  if (!normalized) {
    const error = new Error("A valid timetable date is required.");
    error.statusCode = 400;
    throw error;
  }
  return normalized;
}

function formatWorkbookNameFromIsoDate(isoDate) {
  const [year, month, day] = isoDate.split("-");
  return `${day}_${month}_${year}.xlsx`;
}

function normalizeRecordDate(record) {
  return record.dataDate || record.metrics?.meta?.dataDate || toDateOnlyIso(record.uploadedAt);
}

function toDateOnlyIso(dateValue) {
  const date = new Date(dateValue);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return date.toISOString().slice(0, 10);
}

function mostCommon(values) {
  if (!values.length) {
    return "";
  }

  const counts = new Map();
  for (const value of values) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }

  return [...counts.entries()].sort((left, right) => right[1] - left[1])[0][0];
}

function toMinutes(timeString) {
  const [hours, minutes] = String(timeString).split(":").map(Number);
  return hours * 60 + minutes;
}

function round(value) {
  return Math.round(value * 100) / 100;
}

function sanitizeRecordForMongo(record) {
  const { _id, ...safeRecord } = record;
  return safeRecord;
}

function stripMongoId(record) {
  if (!record) {
    return record;
  }

  const { _id, ...safeRecord } = record;
  return safeRecord;
}
