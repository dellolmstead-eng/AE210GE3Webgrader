import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";

export function runLandingGearChecks(workbook) {
  const feedback = [];
  let good = true;

  const gear = workbook.sheets.gear;

  const noseRule = asNumber(getCell(gear, "J19"));
  if (noseRule != null && (noseRule < 10 || noseRule > 20)) {
    feedback.push(STRINGS.gear.nose);
    good = false;
  }

  const tipbackUpper = asNumber(getCell(gear, "L19"));
  const tipbackLower = asNumber(getCell(gear, "L20"));
  if (tipbackUpper != null && tipbackLower != null && !(tipbackUpper < tipbackLower)) {
    feedback.push(STRINGS.gear.tipback);
    good = false;
  }

  const rolloverUpper = asNumber(getCell(gear, "M19"));
  const rolloverLower = asNumber(getCell(gear, "M20"));
  if (rolloverUpper != null && rolloverLower != null && !(rolloverUpper < rolloverLower)) {
    feedback.push(STRINGS.gear.rollover);
    good = false;
  }

  const rotationSpeed = asNumber(getCell(gear, "N19"));
  if (rotationSpeed != null && !(rotationSpeed < 200)) {
    feedback.push(STRINGS.gear.rotation);
    good = false;
  }

  let delta = 0;
  if (!good) {
    feedback.push(STRINGS.gear.deduction);
    delta -= 1;
  }

  return { delta, feedback };
}
