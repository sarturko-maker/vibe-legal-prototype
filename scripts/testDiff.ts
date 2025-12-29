import { calculateRedline } from '../src/utils/DiffEngine';

const testCase1Original = "The cat sat on the mat";
const testCase1Modified = "The dog sat on the blue mat";

console.log("Test Case 1:");
console.log(calculateRedline(testCase1Original, testCase1Modified));

const testCase2Original = "Governing Law: New York";
const testCase2Modified = "Governing Law: Delaware";

console.log("\nTest Case 2:");
console.log(calculateRedline(testCase2Original, testCase2Modified));
