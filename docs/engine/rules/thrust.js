import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const THRUST_CELLS = [
  ["C48", "C49"],
  ["D48", "D49"],
  ["E48", "E49"],
  ["F48", "F49"],
  ["G48", "G49"],
  ["H48", "H49"],
  ["I48", "I49"],
  ["J48", "J49"],
  ["K48", "K49"],
  ["L48", "L49"],
  ["M48", "M49"],
  ["N48", "N49"],
];

export function runThrustAndTakeoff(workbook) {
  const feedback = [];
  let delta = 0;

  const miss = workbook.sheets.miss;
  const main = workbook.sheets.main;

  let thrustFailures = 0;
  THRUST_CELLS.forEach(([dragRef, availableRef]) => {
    const drag = asNumber(getCell(miss, dragRef));
    const available = asNumber(getCell(miss, availableRef));
    if (!Number.isFinite(drag) || !Number.isFinite(available)) {
      return;
    }
    if (available <= drag) {
      thrustFailures += 1;
    }
  });

  if (thrustFailures > 0) {
    delta -= 1;
    feedback.push(format(STRINGS.thrustLeg, thrustFailures));
    return { delta, feedback };
  }

  const takeoffDistance = asNumber(getCell(main, "K38"));
  const takeoffRequired = asNumber(getCell(main, "X12"));
  if (
    takeoffDistance != null &&
    takeoffRequired != null &&
    takeoffDistance > takeoffRequired
  ) {
    delta -= 1;
    feedback.push(STRINGS.takeoffRoll);
  }

  return { delta, feedback };
}
