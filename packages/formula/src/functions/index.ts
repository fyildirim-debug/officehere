/** Tüm fonksiyon ailelerini kaydeder. */
import { arrayFunctions } from "./array";
import { dateFunctions } from "./date";
import { databaseFunctions } from "./database";
import { distributionFunctions } from "./distributions";
import { engineeringFunctions } from "./engineering";
import { financialFunctions } from "./financial";
import { infoFunctions, logicalFunctions } from "./logical";
import { lookupFunctions } from "./lookup";
import { mathFunctions } from "./math";
import { register } from "./registry";
import { regexFunctions } from "./regex";
import { regressionFunctions } from "./regression";
import { statsFunctions } from "./stats";
import { textFunctions } from "./text";

let done = false;

export function registerAll(): void {
  if (done) return;
  done = true;
  register(mathFunctions);
  register(statsFunctions);
  register(distributionFunctions);
  register(logicalFunctions);
  register(infoFunctions);
  register(textFunctions);
  register(lookupFunctions);
  register(arrayFunctions);
  register(dateFunctions);
  register(financialFunctions);
  register(databaseFunctions);
  register(engineeringFunctions);
  register(regexFunctions);
  register(regressionFunctions);
}

registerAll();

export * from "./registry";
