import { Router } from "express";
import multer from "multer";
import {
  applyOptimization,
  getHistory,
  getHistoryRecord,
  getMetrics,
  getTimetableEditorTemplate,
  optimizeRooms,
  getPrediction,
  saveTimetableDay,
  simulateBatchSize,
  uploadWorkbook
} from "../controllers/metricsController.js";

const router = Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }
});

router.post("/upload", upload.single("file"), uploadWorkbook);
router.get("/metrics", getMetrics);
router.get("/history", getHistory);
router.get("/history/:recordId", getHistoryRecord);
router.get("/predict", getPrediction);
router.post("/simulate", simulateBatchSize);
router.get("/optimize", optimizeRooms);
router.post("/optimize/apply", applyOptimization);
router.get("/timetable/template", getTimetableEditorTemplate);
router.post("/timetable/save", saveTimetableDay);

export default router;
