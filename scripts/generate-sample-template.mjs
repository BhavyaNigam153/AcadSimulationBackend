import ExcelJS from "exceljs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const outputPath = path.resolve(__dirname, "../../academic-block-sample-template.xlsx");

const workbook = new ExcelJS.Workbook();
workbook.creator = "Codex";
workbook.created = new Date();

const rooms = workbook.addWorksheet("Rooms");
rooms.columns = [
  { header: "room_id", key: "room_id", width: 16 },
  { header: "capacity", key: "capacity", width: 12 },
  { header: "type", key: "type", width: 18 },
  { header: "floor", key: "floor", width: 10 }
];
rooms.addRows([
  { room_id: "A101", capacity: 60, type: "Lecture", floor: 1 },
  { room_id: "A102", capacity: 40, type: "Lab", floor: 1 },
  { room_id: "B201", capacity: 120, type: "Seminar", floor: 2 },
  { room_id: "C301", capacity: 80, type: "Lecture", floor: 3 },
  { room_id: "D401", capacity: 35, type: "Tutorial", floor: 4 }
]);

const timetable = workbook.addWorksheet("Timetable");
timetable.columns = [
  { header: "class_id", key: "class_id", width: 16 },
  { header: "subject", key: "subject", width: 24 },
  { header: "room_id", key: "room_id", width: 16 },
  { header: "start_time", key: "start_time", width: 14 },
  { header: "end_time", key: "end_time", width: 14 },
  { header: "day", key: "day", width: 16 }
];
timetable.addRows([
  { class_id: "CSE101", subject: "Data Structures", room_id: "A101", start_time: "09:00", end_time: "10:00", day: "Monday" },
  { class_id: "CSE102", subject: "Database Systems", room_id: "A102", start_time: "10:00", end_time: "11:30", day: "Monday" },
  { class_id: "MAT201", subject: "Calculus II", room_id: "B201", start_time: "09:30", end_time: "10:30", day: "Tuesday" },
  { class_id: "PHY110", subject: "Applied Physics", room_id: "C301", start_time: "11:00", end_time: "12:00", day: "Tuesday" },
  { class_id: "ENG205", subject: "Technical Writing", room_id: "D401", start_time: "14:00", end_time: "15:00", day: "Wednesday" },
  { class_id: "CSE220", subject: "Operating Systems", room_id: "A101", start_time: "13:00", end_time: "14:30", day: "Thursday" },
  { class_id: "ECE210", subject: "Digital Logic", room_id: "C301", start_time: "15:00", end_time: "16:00", day: "Friday" }
]);

const enrollment = workbook.addWorksheet("Enrollment");
enrollment.columns = [
  { header: "class_id", key: "class_id", width: 16 },
  { header: "student_count", key: "student_count", width: 16 }
];
enrollment.addRows([
  { class_id: "CSE101", student_count: 55 },
  { class_id: "CSE102", student_count: 38 },
  { class_id: "MAT201", student_count: 130 },
  { class_id: "PHY110", student_count: 72 },
  { class_id: "ENG205", student_count: 22 },
  { class_id: "CSE220", student_count: 64 },
  { class_id: "ECE210", student_count: 48 }
]);

for (const sheet of [rooms, timetable, enrollment]) {
  sheet.getRow(1).font = { bold: true };
  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

await workbook.xlsx.writeFile(outputPath);
console.log(outputPath);
