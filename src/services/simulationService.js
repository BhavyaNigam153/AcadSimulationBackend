const UPPER_FLOORS = new Set(["Second", "Third"]);
const LOWER_FLOORS = new Set(["Ground", "First"]);
const MBA_DURATION_MINUTES = 90;
const IPM_DURATION_MINUTES = 60;
const MIN_CONSOLIDATION_GAP_MINUTES = 15;

export function runBatchSizeSimulation(metrics, percentageIncrease) {
  if (!metrics?.classes?.length || !metrics?.rooms?.length) {
    const error = new Error("Simulation requires uploaded class-level data.");
    error.statusCode = 400;
    throw error;
  }

  if (!Number.isFinite(percentageIncrease) || percentageIncrease <= 0) {
    const error = new Error("percentageIncrease must be a positive number.");
    error.statusCode = 400;
    throw error;
  }

  const multiplier = 1 + percentageIncrease / 100;
  const simulatedClasses = metrics.classes.map((classItem) => {
    const simulatedStudentCount = Math.round(classItem.studentCount * multiplier);
    const simulatedOccupancy = classItem.capacity
      ? simulatedStudentCount / classItem.capacity
      : 0;

    return {
      classId: classItem.classId,
      subject: classItem.subject,
      roomId: classItem.roomId,
      roomNameEn: classItem.roomNameEn || "",
      program: classItem.program || "Management",
      day: classItem.day,
      startTime: classItem.startTime,
      endTime: classItem.endTime,
      simulatedStudentCount,
      simulatedOccupancy: round(simulatedOccupancy),
      overloadBy: Math.max(simulatedStudentCount - classItem.capacity, 0),
    };
  });

  const avgOccupancy =
    simulatedClasses.reduce((sum, classItem) => sum + classItem.simulatedOccupancy, 0) /
    Math.max(simulatedClasses.length, 1);

  const overloadedRooms = metrics.rooms
    .map((room) => {
      const impactedClasses = simulatedClasses.filter(
        (classItem) =>
          classItem.roomId === room.roomId && classItem.simulatedOccupancy > 1,
      );

      if (!impactedClasses.length) {
        return null;
      }

      return {
        roomId: room.roomId,
        roomNameEn: room.roomNameEn || "",
        floor: room.floor || "",
        overloadedClasses: impactedClasses.length,
        maxOccupancy: round(
          Math.max(...impactedClasses.map((classItem) => classItem.simulatedOccupancy)),
        ),
      };
    })
    .filter(Boolean);

  const criticalClasses = simulatedClasses
    .filter((classItem) => classItem.simulatedOccupancy >= 1)
    .sort((left, right) => right.simulatedOccupancy - left.simulatedOccupancy)
    .slice(0, 12);

  return {
    percentageIncrease,
    avgOccupancy: round(avgOccupancy),
    overloadedRooms,
    criticalClasses,
  };
}

export function optimizeRoomAllocation(metrics) {
  if (!metrics?.classes?.length || !metrics?.rooms?.length) {
    const error = new Error("Optimization requires uploaded class-level data.");
    error.statusCode = 400;
    throw error;
  }

  const state = createOptimizationState(metrics);
  const consolidations = consolidateMbaClasses(state);
  const reallocations = reallocateUpperFloorIpmClasses(state);

  return {
    optimizations: [...consolidations, ...reallocations],
  };
}

function createOptimizationState(metrics) {
  const roomsById = new Map(metrics.rooms.map((room) => [room.roomId, room]));
  const classes = metrics.classes
    .map((classItem) => buildOptimizableClass(classItem, roomsById))
    .filter(Boolean);

  return {
    rooms: metrics.rooms,
    roomsById,
    classes,
    schedulesByRoomId: buildRoomSchedules(metrics.rooms, classes),
    freedLowerFloorRooms: new Set(),
    movedClassIds: new Set(),
  };
}

function buildOptimizableClass(classItem, roomsById) {
  const room = roomsById.get(classItem.roomId);
  if (!room) {
    return null;
  }

  const startMinutes = toMinutes(classItem.startTime);
  const endMinutes = toMinutes(classItem.endTime);

  return {
    ...classItem,
    sessionKey: buildSessionKey(classItem),
    roomId: room.roomId,
    roomNameEn: room.roomNameEn || classItem.roomNameEn || "",
    floor: room.floor || classItem.floor || "",
    type: room.type || classItem.type || "",
    capacity: room.capacity || classItem.capacity || 0,
    startMinutes,
    endMinutes,
    durationMinutes: endMinutes - startMinutes,
  };
}

function buildRoomSchedules(rooms, classes) {
  const schedulesByRoomId = new Map(rooms.map((room) => [room.roomId, new Map()]));

  for (const classItem of classes) {
    addClassToSchedule(schedulesByRoomId, classItem.roomId, classItem);
  }

  return schedulesByRoomId;
}

function consolidateMbaClasses(state) {
  const mbaCandidates = state.classes
    .filter((classItem) => classItem.durationMinutes === MBA_DURATION_MINUTES)
    .filter((classItem) => !state.movedClassIds.has(classItem.classId))
    .sort((left, right) => {
      const floorScore = floorRank(right.floor) - floorRank(left.floor);
      if (floorScore !== 0) {
        return floorScore;
      }

      if (left.day !== right.day) {
        return left.day.localeCompare(right.day);
      }

      return left.startMinutes - right.startMinutes;
    });

  const optimizations = [];

  for (const classItem of mbaCandidates) {
    const targetRoom = findBestConsolidationRoom(classItem, state);
    if (!targetRoom) {
      continue;
    }

    const sourceRoomId = classItem.roomId;
    const sourceRoomName = classItem.roomNameEn || sourceRoomId;

    moveClassToRoom(classItem, targetRoom, state);
    state.movedClassIds.add(classItem.classId);

    if (LOWER_FLOORS.has(getRoomFloor(state.roomsById.get(sourceRoomId)))) {
      state.freedLowerFloorRooms.add(sourceRoomId);
    }

    optimizations.push({
      type: "consolidation",
      classId: classItem.classId,
      day: classItem.day,
      startTime: classItem.startTime,
      endTime: classItem.endTime,
      fromRoom: sourceRoomId,
      toRoom: targetRoom.roomId,
      time: `${classItem.day} ${classItem.startTime}-${classItem.endTime}`,
      benefit: `Consolidates ${classItem.subject} from ${sourceRoomName} into ${targetRoom.roomNameEn || targetRoom.roomId}.`,
    });
  }

  return optimizations;
}

function reallocateUpperFloorIpmClasses(state) {
  const ipmCandidates = state.classes
    .filter((classItem) => classItem.durationMinutes === IPM_DURATION_MINUTES)
    .filter((classItem) => UPPER_FLOORS.has(classItem.floor))
    .filter((classItem) => !state.movedClassIds.has(classItem.classId))
    .sort((left, right) => {
      const floorScore = floorRank(right.floor) - floorRank(left.floor);
      if (floorScore !== 0) {
        return floorScore;
      }

      if (left.day !== right.day) {
        return left.day.localeCompare(right.day);
      }

      return left.startMinutes - right.startMinutes;
    });

  const optimizations = [];

  for (const classItem of ipmCandidates) {
    const targetRoom = findBestReallocationRoom(classItem, state);
    if (!targetRoom) {
      continue;
    }

    const sourceFloor = classItem.floor;
    const sourceRoomId = classItem.roomId;

    moveClassToRoom(classItem, targetRoom, state);
    state.movedClassIds.add(classItem.classId);

    optimizations.push({
      type: "reallocation",
      classId: classItem.classId,
      day: classItem.day,
      startTime: classItem.startTime,
      endTime: classItem.endTime,
      fromRoom: sourceRoomId,
      toRoom: targetRoom.roomId,
      time: `${classItem.day} ${classItem.startTime}-${classItem.endTime}`,
      benefit: `Moves ${classItem.subject} from ${sourceFloor} floor to ${targetRoom.floor} floor with a better lower-floor allocation.`,
    });
  }

  return optimizations;
}

function findBestConsolidationRoom(classItem, state) {
  const candidateRooms = state.rooms
    .filter((room) => LOWER_FLOORS.has(room.floor))
    .filter((room) => room.roomId !== classItem.roomId)
    .filter((room) => room.capacity >= classItem.studentCount)
    .filter((room) =>
      canPlaceClassInRoom(
        room.roomId,
        classItem,
        state.schedulesByRoomId,
        MIN_CONSOLIDATION_GAP_MINUTES,
      ),
    )
    .sort((left, right) => compareConsolidationRooms(left, right, classItem, state));

  return candidateRooms[0] || null;
}

function findBestReallocationRoom(classItem, state) {
  const candidateRooms = state.rooms
    .filter((room) => LOWER_FLOORS.has(room.floor))
    .filter((room) => room.roomId !== classItem.roomId)
    .filter((room) => room.capacity >= classItem.studentCount)
    .filter((room) => canPlaceClassInRoom(room.roomId, classItem, state.schedulesByRoomId, 0))
    .sort((left, right) => compareReallocationRooms(left, right, classItem, state));

  return candidateRooms[0] || null;
}

function compareConsolidationRooms(left, right, classItem, state) {
  const leftDayLoad = getDaySchedule(state.schedulesByRoomId, left.roomId, classItem.day).length;
  const rightDayLoad = getDaySchedule(state.schedulesByRoomId, right.roomId, classItem.day).length;

  if (leftDayLoad !== rightDayLoad) {
    return rightDayLoad - leftDayLoad;
  }

  const leftFloorRank = floorRank(left.floor);
  const rightFloorRank = floorRank(right.floor);
  if (leftFloorRank !== rightFloorRank) {
    return leftFloorRank - rightFloorRank;
  }

  return Math.abs(left.capacity - classItem.studentCount) -
    Math.abs(right.capacity - classItem.studentCount);
}

function compareReallocationRooms(left, right, classItem, state) {
  const leftFreedScore = state.freedLowerFloorRooms.has(left.roomId) ? 0 : 1;
  const rightFreedScore = state.freedLowerFloorRooms.has(right.roomId) ? 0 : 1;

  if (leftFreedScore !== rightFreedScore) {
    return leftFreedScore - rightFreedScore;
  }

  const leftFloorRank = floorRank(left.floor);
  const rightFloorRank = floorRank(right.floor);
  if (leftFloorRank !== rightFloorRank) {
    return leftFloorRank - rightFloorRank;
  }

  const leftDayLoad = getDaySchedule(state.schedulesByRoomId, left.roomId, classItem.day).length;
  const rightDayLoad = getDaySchedule(state.schedulesByRoomId, right.roomId, classItem.day).length;
  if (leftDayLoad !== rightDayLoad) {
    return leftDayLoad - rightDayLoad;
  }

  return Math.abs(left.capacity - classItem.studentCount) -
    Math.abs(right.capacity - classItem.studentCount);
}

function moveClassToRoom(classItem, targetRoom, state) {
  removeClassFromSchedule(state.schedulesByRoomId, classItem.roomId, classItem);

  classItem.roomId = targetRoom.roomId;
  classItem.roomNameEn = targetRoom.roomNameEn || classItem.roomNameEn || "";
  classItem.floor = targetRoom.floor || classItem.floor || "";
  classItem.type = targetRoom.type || classItem.type || "";
  classItem.capacity = targetRoom.capacity || classItem.capacity || 0;

  addClassToSchedule(state.schedulesByRoomId, targetRoom.roomId, classItem);
}

function canPlaceClassInRoom(roomId, classItem, schedulesByRoomId, minimumGapMinutes) {
  const daySchedule = getDaySchedule(schedulesByRoomId, roomId, classItem.day)
    .filter((scheduledClass) => scheduledClass.sessionKey !== classItem.sessionKey);

  let insertAt = daySchedule.length;

  for (let index = 0; index < daySchedule.length; index += 1) {
    const scheduledClass = daySchedule[index];

    if (
      classItem.startMinutes < scheduledClass.endMinutes &&
      classItem.endMinutes > scheduledClass.startMinutes
    ) {
      return false;
    }

    if (classItem.endMinutes <= scheduledClass.startMinutes) {
      insertAt = index;
      break;
    }
  }

  const previousClass = insertAt > 0 ? daySchedule[insertAt - 1] : null;
  const nextClass = insertAt < daySchedule.length ? daySchedule[insertAt] : null;

  const previousGap = previousClass
    ? classItem.startMinutes - previousClass.endMinutes
    : Number.POSITIVE_INFINITY;
  const nextGap = nextClass
    ? nextClass.startMinutes - classItem.endMinutes
    : Number.POSITIVE_INFINITY;

  return previousGap >= minimumGapMinutes && nextGap >= minimumGapMinutes;
}

function addClassToSchedule(schedulesByRoomId, roomId, classItem) {
  const roomSchedule = schedulesByRoomId.get(roomId) || new Map();
  const daySchedule = roomSchedule.get(classItem.day) || [];
  daySchedule.push(classItem);
  daySchedule.sort((left, right) => left.startMinutes - right.startMinutes);
  roomSchedule.set(classItem.day, daySchedule);
  schedulesByRoomId.set(roomId, roomSchedule);
}

function removeClassFromSchedule(schedulesByRoomId, roomId, classItem) {
  const roomSchedule = schedulesByRoomId.get(roomId);
  if (!roomSchedule) {
    return;
  }

  const daySchedule = roomSchedule.get(classItem.day) || [];
  roomSchedule.set(
    classItem.day,
    daySchedule.filter((scheduledClass) => scheduledClass.sessionKey !== classItem.sessionKey),
  );
}

function buildSessionKey(classItem) {
  return [
    classItem.classId,
    classItem.day,
    classItem.startTime,
    classItem.endTime,
    classItem.roomId,
  ].join("|");
}

function getDaySchedule(schedulesByRoomId, roomId, day) {
  return schedulesByRoomId.get(roomId)?.get(day) || [];
}

function getRoomFloor(room) {
  return room?.floor || "";
}

function floorRank(floor) {
  if (floor === "Ground") {
    return 0;
  }
  if (floor === "First") {
    return 1;
  }
  if (floor === "Second") {
    return 2;
  }
  return 3;
}

function toMinutes(timeString) {
  const [hours, minutes] = String(timeString).split(":").map(Number);
  return hours * 60 + minutes;
}

function round(value) {
  return Math.round(value * 100) / 100;
}
