import {
  buildDatasetFromWorkbook,
  buildMetricsFromDataset,
} from "../services/excelService.js";
import {
  optimizeRoomAllocation,
  runBatchSizeSimulation
} from "../services/simulationService.js";
import {
  applyOptimizationToRecord,
  buildPrediction,
  getTimetableTemplate,
  getStoredMetrics,
  getStoredRecord,
  listStoredHistory,
  saveTimetableRecord,
  saveStoredWorkbook
} from "../store/metricsStore.js";

export async function uploadWorkbook(req, res, next) {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Excel file is required." });
    }

    const fileName = req.file.originalname || "";
    if (!/^\d{2}_\d{2}_\d{4}\.xlsx$/i.test(fileName)) {
      return res.status(400).json({
        message: "File name must follow dd_mm_yyyy.xlsx format, for example 29_03_2026.xlsx."
      });
    }

    const dataset = await buildDatasetFromWorkbook(req.file.buffer);
    const metrics = buildMetricsFromDataset(dataset);
    const storedMetrics = await saveStoredWorkbook({
      originalName: req.file.originalname,
      buffer: req.file.buffer,
      dataset,
      metrics,
    });

    return res.status(200).json({
      message: "Excel file uploaded and processed successfully.",
      metrics: storedMetrics
    });
  } catch (error) {
    return next(error);
  }
}

export async function getMetrics(_req, res, next) {
  try {
    const metrics = await getStoredMetrics();

    if (!metrics) {
      return res.status(404).json({
        message: "No metrics available. Upload an Excel file first."
      });
    }

    return res.status(200).json(metrics);
  } catch (error) {
    return next(error);
  }
}

export async function getHistory(_req, res, next) {
  try {
    const history = await listStoredHistory();
    return res.status(200).json(history);
  } catch (error) {
    return next(error);
  }
}

export async function getHistoryRecord(req, res, next) {
  try {
    const metrics = await getStoredRecord(req.params.recordId);

    if (!metrics) {
      return res.status(404).json({
        message: "Requested historical record was not found."
      });
    }

    return res.status(200).json(metrics);
  } catch (error) {
    return next(error);
  }
}

export async function getPrediction(req, res, next) {
  try {
    const daysAhead = Number.parseInt(req.query.daysAhead || "1", 10);

    if (!Number.isInteger(daysAhead) || daysAhead < 1 || daysAhead > 30) {
      return res.status(400).json({
        message: "daysAhead must be an integer between 1 and 30."
      });
    }

    const prediction = await buildPrediction(daysAhead);

    if (!prediction) {
      return res.status(404).json({
        message: "Not enough historical uploads to generate a prediction yet."
      });
    }

    return res.status(200).json(prediction);
  } catch (error) {
    return next(error);
  }
}

export async function simulateBatchSize(req, res, next) {
  try {
    const metrics = await getStoredMetrics();
    if (!metrics) {
      return res.status(404).json({
        message: "No metrics available. Upload an Excel file first."
      });
    }

    const percentageIncrease = Number(req.body?.percentageIncrease);
    const simulation = runBatchSizeSimulation(metrics, percentageIncrease);
    return res.status(200).json(simulation);
  } catch (error) {
    return next(error);
  }
}

export async function optimizeRooms(_req, res, next) {
  try {
    const metrics = _req.query.recordId
      ? await getStoredRecord(_req.query.recordId)
      : await getStoredMetrics();
    if (!metrics) {
      return res.status(404).json({
        message: "No metrics available. Upload an Excel file first."
      });
    }

    const optimization = optimizeRoomAllocation(metrics);
    return res.status(200).json(optimization);
  } catch (error) {
    return next(error);
  }
}

export async function getTimetableEditorTemplate(req, res, next) {
  try {
    const template = await getTimetableTemplate(req.query.recordId || null);
    return res.status(200).json(template);
  } catch (error) {
    return next(error);
  }
}

export async function saveTimetableDay(req, res, next) {
  try {
    const metrics = await saveTimetableRecord({
      dataDate: req.body?.dataDate,
      timetable: req.body?.timetable || [],
      basedOnRecordId: req.body?.basedOnRecordId || null,
      targetRecordId: req.body?.targetRecordId || null,
    });

    return res.status(200).json({
      message: "Timetable saved successfully.",
      metrics,
    });
  } catch (error) {
    return next(error);
  }
}

export async function applyOptimization(req, res, next) {
  try {
    const metrics = await applyOptimizationToRecord({
      recordId: req.body?.recordId || null,
      classId: req.body?.classId,
      day: req.body?.day || null,
      startTime: req.body?.startTime || null,
      toRoom: req.body?.toRoom,
    });

    return res.status(200).json({
      message: "Optimization applied successfully.",
      metrics,
    });
  } catch (error) {
    return next(error);
  }
}
