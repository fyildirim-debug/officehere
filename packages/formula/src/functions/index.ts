/** Tüm fonksiyon ailelerini kaydeder. */
import { arrayFunctions } from "./array";
import { dateFunctions } from "./date";
import { financialFunctions } from "./financial";
import { infoFunctions, logicalFunctions } from "./logical";
import { lookupFunctions } from "./lookup";
import { mathFunctions } from "./math";
import { register } from "./registry";
import { statsFunctions } from "./stats";
import { textFunctions } from "./text";

let done = false;

export function registerAll(): void {
  if (done) return;
  done = true;
  register(mathFunctions);
  register(statsFunctions);
  register(logicalFunctions);
  register(infoFunctions);
  register(textFunctions);
  register(lookupFunctions);
  register(arrayFunctions);
  register(dateFunctions);
  register(financialFunctions);
}

registerAll();

export * from "./registry";
