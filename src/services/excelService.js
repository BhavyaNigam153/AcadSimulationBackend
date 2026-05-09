import ExcelJS from "exceljs";

const REQUIRED_SHEETS = ["Rooms", "Timetable", "Enrollment"];
const BBA_SUBJECTS = new Set([
  "Principles of Management",
  "Business Economics",
  "Financial Accounting",
  "Marketing Fundamentals",
  "Business Statistics",
  "Organizational Behavior",
  "Business Communication",
]);
const MBA_SUBJECTS = new Set([
  "Managerial Economics",
  "Financial Management",
  "Marketing Management",
  "Operations Management",
  "Human Resource Management",
  "Business Analytics",
  "Strategic Management",
  "Corporate Finance",
]);
const UNASSIGNED_ROOM_ID = "TBA";

export async function buildDatasetFromWorkbook(fileBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);

  const roomsSheet = requireSheet(workbook, "Rooms");
  const timetableSheet = requireSheet(workbook, "Timetable");
  const enrollmentSheet = requireSheet(workbook, "Enrollment");

  const rooms = parseRoomsSheet(roomsSheet);
  const timetable = parseTimetableSheet(timetableSheet);
  const enrollments = parseEnrollmentSheet(enrollmentSheet);

  return { rooms, timetable, enrollments };
}

export async function buildMetricsFromWorkbook(fileBuffer) {
  const dataset = await buildDatasetFromWorkbook(fileBuffer);
  return buildMetricsFromDataset(dataset);
}

export function buildMetricsFromDataset({ rooms, timetable, enrollments }) {
  return computeMetrics({ rooms, timetable, enrollments });
}

function requireSheet(workbook, sheetName) {
  const sheet = workbook.getWorksheet(sheetName);

  if (!sheet) {
    throw createBadRequest(
      `Missing required sheet: ${sheetName}. Expected sheets: ${REQUIRED_SHEETS.join(", ")}`,
    );
  }

  return sheet;
}

function parseRoomsSheet(sheet) {
  const rows = sheetToJson(sheet);

  ensureHeaders(
    rows.headers,
    ["room_id", "capacity", "type", "floor"],
    "Rooms",
  );

  if (rows.data.length === 0) {
    throw createBadRequest("Rooms sheet does not contain any data rows.");
  }

  return rows.data.map((row, index) => ({
    roomId: requireString(row.room_id, "room_id", index),
    roomNameEn: requireString(row.room_name_en, "room_name_en", index),
    roomNameHi: requireString(row.room_name_hi, "room_name_hi", index),
    capacity: requireNumber(row.capacity, "capacity", index),
    type: requireString(row.type, "type", index),
    floor: requireString(row.floor, "floor", index),
  }));
}

function parseTimetableSheet(sheet) {
  const rows = sheetToJson(sheet);

  ensureHeaders(
    rows.headers,
    ["class_id", "subject", "room_id", "start_time", "end_time", "day"],
    "Timetable",
  );

  if (rows.data.length === 0) {
    throw createBadRequest("Timetable sheet does not contain any data rows.");
  }

  return rows.data.map((row, index) => {
    const classId = requireString(row.class_id, "class_id", index);
    const subject = requireString(row.subject, "subject", index);
    const roomId = requireString(row.room_id, "room_id", index);
    const day = requireString(row.day, "day", index);
    const startTime = parseTime(row.start_time, "start_time", index);
    const endTime = parseTime(row.end_time, "end_time", index);

    if (toMinutes(endTime) <= toMinutes(startTime)) {
      throw createBadRequest(
        `end_time (${endTime}) must be after start_time (${startTime}) in Timetable row ${index + 2} for ${classId} - ${subject} in room ${roomId} on ${day}.`,
      );
    }

    return {
      classId,
      subject,
      roomId,
      startTime,
      endTime,
      day,
    };
  });
}

function parseEnrollmentSheet(sheet) {
  const rows = sheetToJson(sheet);

  ensureHeaders(rows.headers, ["class_id", "student_count"], "Enrollment");

  if (rows.data.length === 0) {
    throw createBadRequest("Enrollment sheet does not contain any data rows.");
  }

  return rows.data.map((row, index) => ({
    classId: requireString(row.class_id, "class_id", index),
    studentCount: requireNumber(row.student_count, "student_count", index),
  }));
}

function sheetToJson(sheet) {
  const rawRows = [];

  sheet.eachRow({ includeEmpty: false }, (row) => {
    rawRows.push(row.values.slice(1));
  });

  if (rawRows.length === 0) {
    throw createBadRequest(`Sheet ${sheet.name} is empty.`);
  }

  const headers = rawRows[0].map((header) => normalizeHeader(header));
  const data = rawRows
    .slice(1)
    .filter((row) =>
      row.some(
        (value) =>
          value !== null && value !== undefined && String(value).trim() !== "",
      ),
    );

  return {
    headers,
    data: data.map((row) =>
      headers.reduce((acc, header, index) => {
        acc[header] = row[index];
        return acc;
      }, {}),
    ),
  };
}

function normalizeHeader(header) {
  return String(header || "")
    .trim()
    .toLowerCase();
}

function ensureHeaders(headers, expectedHeaders, sheetName) {
  const missing = expectedHeaders.filter((header) => !headers.includes(header));
  if (missing.length > 0) {
    throw createBadRequest(
      `Missing headers in ${sheetName} sheet: ${missing.join(", ")}`,
    );
  }
}

function requireString(value, fieldName, index) {
  const normalized = String(value ?? "").trim();
  if (!normalized) {
    throw createBadRequest(`Missing ${fieldName} in row ${index + 2}.`);
  }
  return normalized;
}

function optionalString(value) {
  return String(value ?? "").trim();
}

function requireNumber(value, fieldName, index) {
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) {
    throw createBadRequest(`Invalid ${fieldName} in row ${index + 2}.`);
  }
  return numericValue;
}

function parseTime(value, fieldName, index) {
  if (value instanceof Date) {
    // Excel time-only cells can arrive as Date objects anchored to 1899.
    // Reading them in local time shifts values because of historical timezone offsets.
    const hours = String(value.getUTCHours()).padStart(2, "0");
    const minutes = String(value.getUTCMinutes()).padStart(2, "0");
    return `${hours}:${minutes}`;
  }

  if (typeof value === "number") {
    const totalMinutes = Math.round(value * 24 * 60);
    const hours = String(Math.floor(totalMinutes / 60) % 24).padStart(2, "0");
    const minutes = String(totalMinutes % 60).padStart(2, "0");
    return `${hours}:${minutes}`;
  }

  const normalized = String(value ?? "").trim();
  if (!/^\d{1,2}:\d{2}$/.test(normalized)) {
    throw createBadRequest(
      `Invalid ${fieldName} in row ${index + 2}. Expected HH:mm.`,
    );
  }

  const [hours, minutes] = normalized.split(":").map(Number);
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
    throw createBadRequest(
      `Invalid ${fieldName} in row ${index + 2}. Expected HH:mm.`,
    );
  }

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function computeMetrics({ rooms, timetable, enrollments }) {
  const roomsById = new Map(rooms.map((room) => [room.roomId, room]));
  const enrollmentsByClassId = new Map(
    enrollments.map((enrollment) => [enrollment.classId, enrollment]),
  );
  const sessionsByRoomId = new Map();
  const timelineEvents = new Map();
  const occupancySamples = [];
  const classes = [];

  let earliestStart = Number.POSITIVE_INFINITY;
  let latestEnd = 0;

  const distinctDays = new Set(
    timetable.map((entry) => String(entry.day).trim()).filter(Boolean),
  );

  for (const entry of timetable) {
    const isUnassignedRoom =
      String(entry.roomId || "").trim().toUpperCase() === UNASSIGNED_ROOM_ID;
    const room = isUnassignedRoom ? null : roomsById.get(entry.roomId);

    if (!room && !isUnassignedRoom) {
      throw createBadRequest(
        `Timetable references unknown room_id: ${entry.roomId}`,
      );
    }

    const enrollment = enrollmentsByClassId.get(entry.classId);
    if (!enrollment) {
      throw createBadRequest(
        `Enrollment is missing class_id: ${entry.classId}`,
      );
    }

    const startMinutes = toMinutes(entry.startTime);
    const endMinutes = toMinutes(entry.endTime);
    earliestStart = Math.min(earliestStart, startMinutes);
    latestEnd = Math.max(latestEnd, endMinutes);

    const occupancy =
      room && room.capacity !== 0 ? enrollment.studentCount / room.capacity : 0;
    const session = {
      ...entry,
      roomNameEn: room?.roomNameEn || (isUnassignedRoom ? UNASSIGNED_ROOM_ID : ""),
      roomNameHi: room?.roomNameHi || "",
      floor: room?.floor || (isUnassignedRoom ? "Unassigned" : ""),
      type: room?.type || "",
      studentCount: enrollment.studentCount,
      occupancy,
      durationMinutes: endMinutes - startMinutes,
      isUnassignedRoom,
    };

    if (room) {
      occupancySamples.push(occupancy);
      const roomSessions = sessionsByRoomId.get(room.roomId) || [];
      roomSessions.push(session);
      sessionsByRoomId.set(room.roomId, roomSessions);
    }

    classes.push(buildClassDetails(session, room));

    timelineEvents.set(
      startMinutes,
      (timelineEvents.get(startMinutes) || 0) + enrollment.studentCount,
    );
    timelineEvents.set(
      endMinutes,
      (timelineEvents.get(endMinutes) || 0) - enrollment.studentCount,
    );
  }

  const availableWindow = Math.max(latestEnd - earliestStart, 1);
  const totalAvailableTime = availableWindow * Math.max(distinctDays.size, 1);

  const roomMetrics = rooms
    .map((room) =>
      buildRoomMetrics(
        room,
        sessionsByRoomId.get(room.roomId) || [],
        rooms,
        totalAvailableTime,
      ),
    )
    .sort((a, b) => a.roomId.localeCompare(b.roomId));

  const timeSeries = buildTimeSeries(timelineEvents);
  const peakHourPoint = timeSeries.reduce(
    (peak, point) => (point.students > peak.students ? point : peak),
    { time: "N/A", students: 0 },
  );

  return {
    summary: {
      totalRooms: rooms.length,
      avgOccupancy: round(
        occupancySamples.length
          ? occupancySamples.reduce((sum, value) => sum + value, 0) /
              occupancySamples.length
          : 0,
      ),
      peakHour: peakHourPoint.time,
    },
    rooms: roomMetrics,
    timeSeries,
    classes: classes.sort(compareClassSessions),
  };
}

function buildRoomMetrics(room, sessions, allRooms, totalAvailableTime) {
  const avgOccupancy = sessions.length
    ? sessions.reduce((sum, session) => sum + session.occupancy, 0) /
      sessions.length
    : 0;
  const usedTime = sessions.reduce(
    (sum, session) => sum + session.durationMinutes,
    0,
  );
  const utilization = totalAvailableTime ? usedTime / totalAvailableTime : 0;
  const overcrowded = sessions.some((session) => session.occupancy > 1);

  let status = "normal";
  if (overcrowded) {
    status = "overutilized";
  } else if (utilization < 0.4) {
    status = "underutilized";
  }

  return {
    roomId: room.roomId,
    roomNameEn: room.roomNameEn || "",
    roomNameHi: room.roomNameHi || "",
    floor: room.floor,
    type: room.type,
    capacity: room.capacity,
    avgOccupancy: round(avgOccupancy),
    utilization: round(utilization),
    status,
    recommendation: overcrowded ? recommendRoom(room, sessions, allRooms) : "",
  };
}

function buildClassDetails(session, room) {
  if (session.isUnassignedRoom) {
    return {
      classId: session.classId,
      program: inferProgram(session.subject),
      subject: session.subject,
      day: session.day,
      startTime: session.startTime,
      endTime: session.endTime,
      roomId: session.roomId,
      roomNameEn: session.roomNameEn || UNASSIGNED_ROOM_ID,
      roomNameHi: "",
      floor: "Unassigned",
      type: "",
      capacity: 0,
      studentCount: session.studentCount,
      occupancy: 0,
      status: "unassigned",
    };
  }

  let status = "normal";
  if (session.occupancy > 1) {
    status = "overcrowded";
  } else if (session.occupancy < 0.4) {
    status = "underutilized";
  }

  return {
    classId: session.classId,
    program: inferProgram(session.subject),
    subject: session.subject,
    day: session.day,
    startTime: session.startTime,
    endTime: session.endTime,
    roomId: room.roomId,
    roomNameEn: room.roomNameEn || "",
    roomNameHi: room.roomNameHi || "",
    floor: room.floor,
    type: room.type,
    capacity: room.capacity,
    studentCount: session.studentCount,
    occupancy: round(session.occupancy),
    status,
  };
}

function compareClassSessions(left, right) {
  const dayCompare = left.day.localeCompare(right.day);
  if (dayCompare !== 0) {
    return dayCompare;
  }

  return toMinutes(left.startTime) - toMinutes(right.startTime);
}

function inferProgram(subject) {
  if (BBA_SUBJECTS.has(subject)) {
    return "BBA";
  }

  if (MBA_SUBJECTS.has(subject)) {
    return "MBA";
  }

  return "Management";
}

function recommendRoom(currentRoom, sessions, allRooms) {
  const requiredCapacity = Math.max(
    ...sessions.map((session) => session.studentCount),
    0,
  );

  const candidate = [...allRooms]
    .filter(
      (room) =>
        room.roomId !== currentRoom.roomId && room.capacity >= requiredCapacity,
    )
    .sort((a, b) => a.capacity - b.capacity)[0];

  if (!candidate) {
    return "No larger room currently available in uploaded inventory.";
  }

  const displayName =
    candidate.roomNameEn || candidate.roomNameHi || "another room";
  const hindiLabel = candidate.roomNameHi ? ` / ${candidate.roomNameHi}` : "";
  return `Consider moving to ${displayName}${hindiLabel} (${candidate.floor}, capacity ${candidate.capacity}).`;
}

function buildTimeSeries(timelineEvents) {
  const sortedMinutes = [...timelineEvents.keys()].sort((a, b) => a - b);
  let runningStudents = 0;

  return sortedMinutes.map((minute) => {
    runningStudents += timelineEvents.get(minute) || 0;
    return {
      time: fromMinutes(minute),
      students: Math.max(runningStudents, 0),
    };
  });
}

function toMinutes(timeString) {
  const [hours, minutes] = String(timeString).split(":").map(Number);
  return hours * 60 + minutes;
}

function fromMinutes(totalMinutes) {
  const hours = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
  const minutes = String(totalMinutes % 60).padStart(2, "0");
  return `${hours}:${minutes}`;
}

function round(value) {
  return Math.round(value * 100) / 100;
}

function createBadRequest(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}
