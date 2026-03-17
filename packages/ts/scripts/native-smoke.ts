import { getNativeBindingStatus, nativeAdd, nativeGreeting } from "../src/index";

const status = getNativeBindingStatus();
console.log("status", status.available, status.reason);

if (!status.available) {
  process.exitCode = 1;
} else {
  console.log("add", await nativeAdd(2, 3));
  console.log("greeting", await nativeGreeting());
}
