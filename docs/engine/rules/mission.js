import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";

const LEG_COLUMNS = [11, 12, 13, 14, 16, 18, 19, 22, 23]; // 1-based column indices
const ROWS = {
  altitude: 33,
  mach: 35,
  afterburner: 36,
  distance: 38,
  time: 39,
};
const missionTolAlt = 10;
const missionTolMach = 0.05;
const missionTolTime = 0.1;
const missionTolDist = 0.5;

export function runMissionChecks(workbook) {
  const feedback = [];
  let missionFailed = false;

  const main = workbook.sheets.main;
  const constraintsMach = asNumber(getCell(main, "U4"));

  const readRowValues = (rowIndex) =>
    LEG_COLUMNS.map((col) => asNumber(getCellByIndex(main, rowIndex, col)));

  const altitude = readRowValues(ROWS.altitude);
  const mach = readRowValues(ROWS.mach);
  const afterburner = readRowValues(ROWS.afterburner);
  const distance = readRowValues(ROWS.distance);
  const time = readRowValues(ROWS.time);

  if (Math.abs(altitude[0] - 0) > missionTolAlt || Math.abs(afterburner[0] - 100) > missionTolAlt) {
    feedback.push(STRINGS.missionLegs[0]);
    missionFailed = true;
  }

  if (!(altitude[1] >= altitude[0] - missionTolAlt && altitude[1] <= altitude[2] + missionTolAlt)) {
    feedback.push(STRINGS.missionLegs[1]);
    missionFailed = true;
  }
  if (!(mach[1] >= mach[0] - missionTolMach && mach[1] <= mach[2] + missionTolMach)) {
    feedback.push(STRINGS.missionLegs[2]);
    missionFailed = true;
  }
  if (Math.abs(afterburner[1] - 0) > missionTolAlt) {
    feedback.push(STRINGS.missionLegs[3]);
    missionFailed = true;
  }

  if (altitude[2] < 35000 - missionTolAlt || Math.abs(mach[2] - 0.9) > missionTolMach || Math.abs(afterburner[2] - 0) > missionTolAlt) {
    feedback.push(STRINGS.missionLegs[4]);
    missionFailed = true;
  }

  if (altitude[3] < 35000 - missionTolAlt || Math.abs(mach[3] - 0.9) > missionTolMach || Math.abs(afterburner[3] - 0) > missionTolAlt) {
    feedback.push(STRINGS.missionLegs[5]);
    missionFailed = true;
  }

  if (
    altitude[4] < 35000 - missionTolAlt ||
    constraintsMach == null ||
    Math.abs(mach[4] - constraintsMach) > missionTolMach ||
    Math.abs(afterburner[4] - 0) > missionTolAlt ||
    distance[4] < 150 - missionTolDist
  ) {
    feedback.push(STRINGS.missionLegs[6]);
    missionFailed = true;
  }

  if (altitude[5] < 30000 - missionTolAlt || mach[5] < 1.2 - missionTolMach || Math.abs(afterburner[5] - 100) > missionTolAlt || time[5] < 2 - missionTolTime) {
    feedback.push(STRINGS.missionLegs[7]);
    missionFailed = true;
  }

  if (
    altitude[6] < 35000 - missionTolAlt ||
    constraintsMach == null ||
    Math.abs(mach[6] - constraintsMach) > missionTolMach ||
    Math.abs(afterburner[6] - 0) > missionTolAlt ||
    distance[6] < 150 - missionTolDist
  ) {
    feedback.push(STRINGS.missionLegs[8]);
    missionFailed = true;
  }

  if (altitude[7] < 35000 - missionTolAlt || Math.abs(mach[7] - 0.9) > missionTolMach || Math.abs(afterburner[7] - 0) > missionTolAlt) {
    feedback.push(STRINGS.missionLegs[9]);
    missionFailed = true;
  }

  if (
    Math.abs(altitude[8] - 10000) > missionTolAlt ||
    Math.abs(mach[8] - 0.4) > missionTolMach ||
    Math.abs(afterburner[8] - 0) > missionTolAlt ||
    Math.abs(time[8] - 20) > missionTolTime
  ) {
    feedback.push(STRINGS.missionLegs[10]);
    missionFailed = true;
  }

  if (missionFailed) {
    feedback.push(STRINGS.missionSummary);
  }

  return { delta: 0, feedback };
}
