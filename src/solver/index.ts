import { TimetableRow } from "../parsers";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } from "docx";

export type Assignment = TimetableRow & { assignedSlot: string };

export type Conflict = {
  type: 'teacher' | 'room' | 'group';
  slot: string;
  conflictingEntries: Assignment[];
  resource: string; // teacher name, room name, or group name
};

export type MergeResult = {
  assignments: Assignment[];
  conflicts: Conflict[];
  separatedTimetables: Map<string, Assignment[]>;
};

const DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"];
const PERIODS = ["P1", "P2", "P3", "P4", "P5", "P6"];

function slotFromRow(r: TimetableRow): string {
  const day = (r.day || "").trim();
  const period = (r.period || "").replace(/^period\s*/i, "P").trim();
  
  // Normalize day names (handle common variations)
  const normalizedDay = normalizeDay(day);
  const normalizedPeriod = normalizePeriod(period);
  
  if (!normalizedDay || !normalizedPeriod) return "";
  return `${normalizedDay}-${normalizedPeriod}`;
}

function normalizeDay(day: string): string {
  const dayLower = day.toLowerCase();
  const dayMappings: Record<string, string> = {
    'monday': 'Mon',
    'tuesday': 'Tue', 
    'wednesday': 'Wed',
    'thursday': 'Thu',
    'friday': 'Fri',
    'mon': 'Mon',
    'tue': 'Tue',
    'wed': 'Wed',
    'thu': 'Thu',
    'fri': 'Fri'
  };
  
  return dayMappings[dayLower] || day;
}

function normalizePeriod(period: string): string {
  // Handle various period formats: "Period 1", "P1", "1", etc.
  const periodMatch = period.match(/(\d+)/);
  if (periodMatch) {
    const periodNum = parseInt(periodMatch[1]);
    if (periodNum >= 1 && periodNum <= 6) {
      return `P${periodNum}`;
    }
  }
  
  // If already in correct format, return as is
  if (PERIODS.includes(period)) {
    return period;
  }
  
  return period;
}

function findNextAvailableSlot(
  currentSlot: string,
  teacher: string,
  room: string,
  group: string,
  occupiedSlots: Map<string, { teachers: Set<string>; rooms: Set<string>; groups: Set<string> }>
): string | null {
  if (!currentSlot) return null;
  
  const [currentDay, currentPeriod] = currentSlot.split('-');
  const dayIndex = DAYS.indexOf(currentDay);
  const periodIndex = PERIODS.indexOf(currentPeriod);
  
  if (dayIndex === -1 || periodIndex === -1) return null;

  // Try next periods on same day first
  for (let p = periodIndex + 1; p < PERIODS.length; p++) {
    const slot = `${currentDay}-${PERIODS[p]}`;
    if (isSlotAvailable(slot, teacher, room, group, occupiedSlots)) {
      return slot;
    }
  }

  // Try next days
  for (let d = dayIndex + 1; d < DAYS.length; d++) {
    for (let p = 0; p < PERIODS.length; p++) {
      const slot = `${DAYS[d]}-${PERIODS[p]}`;
      if (isSlotAvailable(slot, teacher, room, group, occupiedSlots)) {
        return slot;
      }
    }
  }

  // Try previous periods on same day
  for (let p = periodIndex - 1; p >= 0; p--) {
    const slot = `${currentDay}-${PERIODS[p]}`;
    if (isSlotAvailable(slot, teacher, room, group, occupiedSlots)) {
      return slot;
    }
  }

  // Try previous days
  for (let d = dayIndex - 1; d >= 0; d--) {
    for (let p = 0; p < PERIODS.length; p++) {
      const slot = `${DAYS[d]}-${PERIODS[p]}`;
      if (isSlotAvailable(slot, teacher, room, group, occupiedSlots)) {
        return slot;
      }
    }
  }

  return null; // No available slot found
}

function isSlotAvailable(
  slot: string,
  teacher: string,
  room: string,
  group: string,
  occupiedSlots: Map<string, { teachers: Set<string>; rooms: Set<string>; groups: Set<string> }>
): boolean {
  const slotData = occupiedSlots.get(slot);
  if (!slotData) return true;

  const teacherConflict = teacher && slotData.teachers.has(teacher);
  const roomConflict = room && slotData.rooms.has(room);
  const groupConflict = group && slotData.groups.has(group);

  return !teacherConflict && !roomConflict && !groupConflict;
}

function addToSlot(
  slot: string,
  teacher: string,
  room: string,
  group: string,
  occupiedSlots: Map<string, { teachers: Set<string>; rooms: Set<string>; groups: Set<string> }>
): void {
  if (!occupiedSlots.has(slot)) {
    occupiedSlots.set(slot, {
      teachers: new Set(),
      rooms: new Set(),
      groups: new Set()
    });
  }
  
  const slotData = occupiedSlots.get(slot)!;
  if (teacher && teacher.trim()) slotData.teachers.add(teacher.trim());
  if (room && room.trim()) slotData.rooms.add(room.trim());
  if (group && group.trim()) slotData.groups.add(group.trim());
}

export function mergeAndResolve(rows: TimetableRow[]): MergeResult {
  const assignments: Assignment[] = [];
  const separatedTimetables = new Map<string, Assignment[]>();
  const occupiedSlots = new Map<string, { teachers: Set<string>; rooms: Set<string>; groups: Set<string> }>();

  // Filter out invalid rows
  const validRows = rows.filter(row => {
    const slot = slotFromRow(row);
    return slot && slot.includes('-') && (row.teacher || row.room || row.group);
  });

  // First pass: assign slots and resolve conflicts
  for (const row of validRows) {
    const originalSlot = slotFromRow(row);
    let assignedSlot = originalSlot;
    
    const teacher = (row.teacher || "").trim();
    const room = (row.room || "").trim();
    const group = (row.group || "").trim();
    const sourceFile = row.sourceFile || "unknown";

    // Check for conflicts in the original slot
    const hasConflict = !isSlotAvailable(originalSlot, teacher, room, group, occupiedSlots);

    if (hasConflict) {
      // Try to find an alternative slot
      const alternativeSlot = findNextAvailableSlot(originalSlot, teacher, room, group, occupiedSlots);
      
      if (alternativeSlot) {
        assignedSlot = alternativeSlot;
        console.log(`Moved ${row.subject} from ${originalSlot} to ${alternativeSlot} due to conflict`);
      } else {
        // No alternative found, keep original slot (this will create a conflict)
        assignedSlot = originalSlot;
        console.warn(`No alternative slot found for ${row.subject} at ${originalSlot}`);
      }
    }

    const assignment: Assignment = { ...row, assignedSlot };
    assignments.push(assignment);

    // Add to separated timetables
    if (!separatedTimetables.has(sourceFile)) {
      separatedTimetables.set(sourceFile, []);
    }
    separatedTimetables.get(sourceFile)!.push(assignment);

    // Update occupied slots
    addToSlot(assignedSlot, teacher, room, group, occupiedSlots);
  }

  // Second pass: detect remaining conflicts after resolution
  const conflicts = detectConflicts(assignments);

  return { assignments, conflicts, separatedTimetables };
}

function detectConflicts(assignments: Assignment[]): Conflict[] {
  const slotMap = new Map<string, Assignment[]>();
  const conflicts: Conflict[] = [];
  
  // Group assignments by their final assigned slot
  for (const assignment of assignments) {
    if (!assignment.assignedSlot) continue;
    
    if (!slotMap.has(assignment.assignedSlot)) {
      slotMap.set(assignment.assignedSlot, []);
    }
    slotMap.get(assignment.assignedSlot)!.push(assignment);
  }

  // Check each slot for conflicts
  for (const [slot, slotAssignments] of slotMap) {
    if (slotAssignments.length <= 1) continue;

    // Check for teacher conflicts
    const teacherMap = new Map<string, Assignment[]>();
    for (const assignment of slotAssignments) {
      const teacher = (assignment.teacher || "").trim();
      if (teacher) {
        if (!teacherMap.has(teacher)) {
          teacherMap.set(teacher, []);
        }
        teacherMap.get(teacher)!.push(assignment);
      }
    }

    for (const [teacher, conflictingEntries] of teacherMap) {
      if (conflictingEntries.length > 1) {
        conflicts.push({
          type: 'teacher',
          slot,
          resource: teacher,
          conflictingEntries
        });
      }
    }

    // Check for room conflicts
    const roomMap = new Map<string, Assignment[]>();
    for (const assignment of slotAssignments) {
      const room = (assignment.room || "").trim();
      if (room) {
        if (!roomMap.has(room)) {
          roomMap.set(room, []);
        }
        roomMap.get(room)!.push(assignment);
      }
    }

    for (const [room, conflictingEntries] of roomMap) {
      if (conflictingEntries.length > 1) {
        conflicts.push({
          type: 'room',
          slot,
          resource: room,
          conflictingEntries
        });
      }
    }

    // Check for group conflicts
    const groupMap = new Map<string, Assignment[]>();
    for (const assignment of slotAssignments) {
      const group = (assignment.group || "").trim();
      if (group) {
        if (!groupMap.has(group)) {
          groupMap.set(group, []);
        }
        groupMap.get(group)!.push(assignment);
      }
    }

    for (const [group, conflictingEntries] of groupMap) {
      if (conflictingEntries.length > 1) {
        conflicts.push({
          type: 'group',
          slot,
          resource: group,
          conflictingEntries
        });
      }
    }
  }

  return conflicts;
}

// Utility function to get a clean timetable view
export function getTimetableView(assignments: Assignment[]): Map<string, Map<string, Assignment[]>> {
  const timetableView = new Map<string, Map<string, Assignment[]>>();
  
  for (const assignment of assignments) {
    if (!assignment.assignedSlot) continue;
    
    const [day, period] = assignment.assignedSlot.split('-');
    
    if (!timetableView.has(day)) {
      timetableView.set(day, new Map());
    }
    
    const dayMap = timetableView.get(day)!;
    if (!dayMap.has(period)) {
      dayMap.set(period, []);
    }
    
    dayMap.get(period)!.push(assignment);
  }
  
  return timetableView;
}

// Enhanced function to display timetable in a readable format
export function displayTimetable(assignments: Assignment[]): string {
  const timetableView = getTimetableView(assignments);
  let output = "TIMETABLE VIEW:\n";
  output += "=".repeat(80) + "\n\n";
  
  // Header
  output += "Day".padEnd(10);
  PERIODS.forEach(period => {
    output += period.padEnd(15);
  });
  output += "\n";
  output += "-".repeat(80) + "\n";
  
  // Content
  DAYS.forEach(day => {
    output += day.padEnd(10);
    const dayMap = timetableView.get(day);
    
    PERIODS.forEach(period => {
      const assignments = dayMap?.get(period) || [];
      let cellContent = "";
      
      if (assignments.length === 0) {
        cellContent = "-";
      } else if (assignments.length === 1) {
        const a = assignments[0];
        cellContent = `${a.subject || 'N/A'}`;
      } else {
        cellContent = `[${assignments.length} items]`;
      }
      
      output += cellContent.substring(0, 14).padEnd(15);
    });
    
    output += "\n";
  });
  
  return output;
}

// Utility function to format conflicts for display
export function formatConflicts(conflicts: Conflict[]): string {
  if (conflicts.length === 0) {
    return "No conflicts detected.";
  }
  
  let output = `Found ${conflicts.length} conflicts:\n\n`;
  
  for (const conflict of conflicts) {
    output += `${conflict.type.toUpperCase()} CONFLICT at ${conflict.slot}:\n`;
    output += `Resource: ${conflict.resource}\n`;
    output += `Conflicting entries:\n`;
    
    for (const entry of conflict.conflictingEntries) {
      output += `  - ${entry.subject || 'Unknown Subject'} (Group: ${entry.group || 'N/A'}) `;
      output += `taught by ${entry.teacher || 'N/A'} in ${entry.room || 'N/A'}\n`;
    }
    output += '\n';
  }
  
  return output;
}

export async function assignmentsToDocx(assignments: Assignment[], filename: string): Promise<Buffer> {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph("Day")] }),
        new TableCell({ children: [new Paragraph("Period")] }),
        new TableCell({ children: [new Paragraph("Assigned Slot")] }),
        new TableCell({ children: [new Paragraph("Subject")] }),
        new TableCell({ children: [new Paragraph("Teacher")] }),
        new TableCell({ children: [new Paragraph("Group")] }),
        new TableCell({ children: [new Paragraph("Room")] }),
      ],
    }),
    ...assignments.map(a => new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(a.day || "")] }),
        new TableCell({ children: [new Paragraph(a.period || "")] }),
        new TableCell({ children: [new Paragraph(a.assignedSlot || "")] }),
        new TableCell({ children: [new Paragraph(a.subject || "")] }),
        new TableCell({ children: [new Paragraph(a.teacher || "")] }),
        new TableCell({ children: [new Paragraph(a.group || "")] }),
        new TableCell({ children: [new Paragraph(a.room || "")] }),
      ],
    }))
  ];

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({ 
            children: [new TextRun({ text: filename, bold: true, size: 24 })]
          }),
          new Paragraph({ text: "" }), // Empty paragraph for spacing
          new Table({ 
            rows: tableRows,
            width: {
              size: 100,
              type: "pct"
            }
          })
        ],
      },
    ],
  });
  
  return await Packer.toBuffer(doc);
}