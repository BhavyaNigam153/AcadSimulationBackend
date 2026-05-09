import ExcelJS from "exceljs";
import path from "node:path";
import { promises as fs } from "node:fs";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const outputDirectory = path.resolve(__dirname, "../../generated_data");
const endDate = new Date("2026-03-29T00:00:00");

const rooms = [
  { room_id: "019", room_name_en: "Aryabhatta", room_name_hi: "आर्यभट्ट", capacity: 120, type: "Lecture Hall", floor: "Ground" },
  { room_id: "021", room_name_en: "Arjuna", room_name_hi: "अर्जुन", capacity: 110, type: "Lecture Hall", floor: "Ground" },
  { room_id: "023", room_name_en: "Chanakya", room_name_hi: "चाणक्य", capacity: 95, type: "Lecture Hall", floor: "Ground" },
  { room_id: "025", room_name_en: "Dronacharya", room_name_hi: "द्रोणाचार्य", capacity: 100, type: "Lecture Hall", floor: "Ground" },
  { room_id: "027", room_name_en: "Prahalad", room_name_hi: "प्रह्लाद", capacity: 80, type: "Meeting Room", floor: "Ground" },
  { room_id: "029", room_name_en: "Bharat", room_name_hi: "भरत", capacity: 70, type: "Lecture Hall", floor: "Ground" },
  { room_id: "031", room_name_en: "Kashyap", room_name_hi: "कश्यप", capacity: 60, type: "Meeting Room", floor: "Ground" },
  { room_id: "119", room_name_en: "Kripacharya", room_name_hi: "कृपाचार्य", capacity: 90, type: "Lecture Hall", floor: "First" },
  { room_id: "121", room_name_en: "Markandey", room_name_hi: "मार्कण्डेय", capacity: 75, type: "Lecture Hall", floor: "First" },
  { room_id: "123", room_name_en: "Nagarjuna", room_name_hi: "नागार्जुन", capacity: 65, type: "Meeting Room", floor: "First" },
  { room_id: "125", room_name_en: "Patanjali", room_name_hi: "पतंजलि", capacity: 60, type: "Meeting Room", floor: "First" },
  { room_id: "127", room_name_en: "Vivekananda", room_name_hi: "विवेकानंद", capacity: 85, type: "Lecture Hall", floor: "First" },
  { room_id: "128", room_name_en: "Shankaracharya", room_name_hi: "शंकराचार्य", capacity: 55, type: "Meeting Room", floor: "First" },
  { room_id: "129", room_name_en: "Valmiki", room_name_hi: "वाल्मीकि", capacity: 50, type: "Meeting Room", floor: "First" },
  { room_id: "130", room_name_en: "Vashishth", room_name_hi: "वशिष्ठ", capacity: 45, type: "Meeting Room", floor: "First" },
  { room_id: "131", room_name_en: "Vyas", room_name_hi: "व्यास", capacity: 40, type: "Meeting Room", floor: "First" },
  { room_id: "219", room_name_en: "Ramanuj", room_name_hi: "रामानुज", capacity: 70, type: "Lecture Hall", floor: "Second" },
  { room_id: "221", room_name_en: "Attri", room_name_hi: "अत्रि", capacity: 55, type: "Meeting Room", floor: "Second" },
  { room_id: "223", room_name_en: "Eklavya", room_name_hi: "एकलव्य", capacity: 50, type: "Meeting Room", floor: "Second" },
  { room_id: "319", room_name_en: "Pulastya", room_name_hi: "पुलस्त्य", capacity: 45, type: "Meeting Room", floor: "Third" },
  { room_id: "321", room_name_en: "Sandipani", room_name_hi: "संदीपनि", capacity: 35, type: "Meeting Room", floor: "Third" }
];

const bbaSubjects = [
  "Principles of Management",
  "Business Economics",
  "Financial Accounting",
  "Marketing Fundamentals",
  "Business Statistics",
  "Organizational Behavior",
  "Business Communication"
];

const mbaSubjects = [
  "Managerial Economics",
  "Financial Management",
  "Marketing Management",
  "Operations Management",
  "Human Resource Management",
  "Business Analytics",
  "Strategic Management",
  "Corporate Finance"
];

const timeSlots = ["09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00"];
const preferredPeakSlots = new Set(["10:00", "11:00", "12:00", "13:00"]);
const weekdayClassTargets = {
  Monday: 34,
  Tuesday: 32,
  Wednesday: 36,
  Thursday: 31,
  Friday: 33,
  Saturday: 24,
  Sunday: 20
};

await fs.mkdir(outputDirectory, { recursive: true });

const summaries = [];

for (let dayOffset = 9; dayOffset >= 0; dayOffset -= 1) {
  const currentDate = new Date(endDate);
  currentDate.setDate(endDate.getDate() - dayOffset);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Codex";
  workbook.created = new Date();

  const roomsSheet = workbook.addWorksheet("Rooms");
  roomsSheet.columns = [
    { header: "room_id", key: "room_id", width: 12 },
    { header: "room_name_en", key: "room_name_en", width: 18 },
    { header: "room_name_hi", key: "room_name_hi", width: 18 },
    { header: "capacity", key: "capacity", width: 12 },
    { header: "type", key: "type", width: 16 },
    { header: "floor", key: "floor", width: 12 }
  ];
  roomsSheet.addRows(rooms);

  const timetableSheet = workbook.addWorksheet("Timetable");
  timetableSheet.columns = [
    { header: "class_id", key: "class_id", width: 18 },
    { header: "subject", key: "subject", width: 28 },
    { header: "room_id", key: "room_id", width: 12 },
    { header: "start_time", key: "start_time", width: 12 },
    { header: "end_time", key: "end_time", width: 12 },
    { header: "day", key: "day", width: 14 }
  ];

  const enrollmentSheet = workbook.addWorksheet("Enrollment");
  enrollmentSheet.columns = [
    { header: "class_id", key: "class_id", width: 18 },
    { header: "student_count", key: "student_count", width: 16 }
  ];

  const seed = Number(formatDateForFile(currentDate).replaceAll("_", ""));
  const random = createSeededRandom(seed);
  const weekday = currentDate.toLocaleDateString("en-US", { weekday: "long" });
  const targetClasses = weekdayClassTargets[weekday] + Math.floor(random() * 3) - 1;

  const roomPreferences = rooms.map((room) => ({
    ...room,
    weight: getRoomWeight(room)
  }));
  const occupancyTracker = new Set();
  const timetableRows = [];
  const enrollmentRows = [];
  let classIndex = 1;

  while (timetableRows.length < targetClasses) {
    const program = random() < 0.62 ? "BBA" : "MBA";
    const subject = pickSubject(program, random);
    const room = pickRoom(program, roomPreferences, random);
    const slot = pickTimeSlot(random);
    const trackerKey = `${room.room_id}-${slot}`;

    if (occupancyTracker.has(trackerKey)) {
      continue;
    }

    occupancyTracker.add(trackerKey);

    const classId = buildClassId(program, currentDate, classIndex);
    const endTime = incrementHour(slot);
    const studentCount = getStudentCount(room.capacity, timetableRows.length, random);

    timetableRows.push({
      class_id: classId,
      subject,
      room_id: room.room_id,
      start_time: slot,
      end_time: endTime,
      day: weekday
    });

    enrollmentRows.push({
      class_id: classId,
      student_count: studentCount
    });

    classIndex += 1;
  }

  timetableSheet.addRows(
    timetableRows.sort((left, right) => {
      if (left.start_time !== right.start_time) {
        return left.start_time.localeCompare(right.start_time);
      }
      return left.room_id.localeCompare(right.room_id);
    })
  );
  enrollmentSheet.addRows(enrollmentRows);

  for (const sheet of [roomsSheet, timetableSheet, enrollmentSheet]) {
    sheet.getRow(1).font = { bold: true };
    sheet.views = [{ state: "frozen", ySplit: 1 }];
  }

  const fileName = `${formatDateForFile(currentDate)}.xlsx`;
  const outputPath = path.join(outputDirectory, fileName);
  await workbook.xlsx.writeFile(outputPath);

  summaries.push({
    fileName,
    classCount: timetableRows.length,
    peakClasses: timetableRows.filter((row) => preferredPeakSlots.has(row.start_time)).length
  });
}

const busiestDay = summaries.reduce((best, current) =>
  current.classCount > best.classCount ? current : best
);

console.log(`total files: ${summaries.length}`);
console.log(
  `total classes: ${summaries.reduce((sum, item) => sum + item.classCount, 0)}`
);
console.log(`busiest day: ${busiestDay.fileName} (${busiestDay.classCount} classes)`);

function getRoomWeight(room) {
  if (room.floor === "Ground") {
    return 1.4;
  }
  if (room.floor === "First") {
    return 1.15;
  }
  if (room.floor === "Second") {
    return 0.9;
  }
  return 0.7;
}

function pickSubject(program, random) {
  const subjects = program === "MBA" ? mbaSubjects : bbaSubjects;
  return subjects[Math.floor(random() * subjects.length)];
}

function pickRoom(program, roomPreferences, random) {
  const preferred = roomPreferences.filter((room) =>
    program === "MBA" ? room.capacity >= 70 : room.capacity >= 40
  );
  return weightedChoice(preferred.length ? preferred : roomPreferences, random);
}

function pickTimeSlot(random) {
  const slotWeights = timeSlots.map((slot) => (preferredPeakSlots.has(slot) ? 2.8 : 1));
  return weightedChoice(
    timeSlots.map((slot, index) => ({ value: slot, weight: slotWeights[index] })),
    random
  );
}

function getStudentCount(capacity, classSequence, random) {
  if (classSequence % 10 === 0) {
    return Math.min(capacity + 6 + Math.floor(random() * 12), Math.round(capacity * 1.18));
  }

  if (classSequence % 7 === 0) {
    return Math.max(12, Math.floor(capacity * (0.22 + random() * 0.15)));
  }

  return Math.max(
    18,
    Math.floor(capacity * (0.6 + random() * 0.4))
  );
}

function buildClassId(program, date, index) {
  const dateStamp = formatDateForFile(date).replaceAll("_", "");
  return `${program}${String(index).padStart(2, "0")}_${dateStamp}`;
}

function incrementHour(timeValue) {
  const [hours, minutes] = timeValue.split(":").map(Number);
  const nextHour = hours + 1;
  return `${String(nextHour).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function formatDateForFile(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}_${month}_${year}`;
}

function weightedChoice(items, random) {
  const normalizedItems = items.map((item) =>
    item.value ? item : { value: item, weight: item.weight || 1 }
  );
  const totalWeight = normalizedItems.reduce((sum, item) => sum + item.weight, 0);
  let threshold = random() * totalWeight;

  for (const item of normalizedItems) {
    threshold -= item.weight;
    if (threshold <= 0) {
      return item.value;
    }
  }

  return normalizedItems[normalizedItems.length - 1].value;
}

function createSeededRandom(seed) {
  let current = seed % 2147483647;
  if (current <= 0) {
    current += 2147483646;
  }

  return function nextRandom() {
    current = (current * 16807) % 2147483647;
    return (current - 1) / 2147483646;
  };
}
