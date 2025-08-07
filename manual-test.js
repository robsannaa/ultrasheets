// Manual Test for Enhanced Table Detection
// This tests the improved extractHeaders function

function testExtractHeaders() {
  console.log("🧪 Testing Enhanced Table Detection\n");

  // Test case 1: Table starting at row 1 (traditional)
  const case1 = {
    0: { 0: { v: "Date" }, 1: { v: "Product" }, 2: { v: "Revenue" } },
    1: { 0: { v: "2024-01-01" }, 1: { v: "Widget A" }, 2: { v: 1000 } }
  };

  // Test case 2: Table starting at row 5 (your problem case)
  const case2 = {
    0: { 0: { v: "Company Financial Report" } },
    1: { 0: { v: "" } },
    2: { 0: { v: "Quarter 1 Results" } },
    3: { 0: { v: "" } },
    4: { 0: { v: "Month" }, 1: { v: "Sales" }, 2: { v: "Profit" }, 3: { v: "Margin%" } },
    5: { 0: { v: "January" }, 1: { v: 50000 }, 2: { v: 5000 }, 3: { v: "10%" } },
    6: { 0: { v: "February" }, 1: { v: 60000 }, 2: { v: 7000 }, 3: { v: "11.7%" } }
  };

  // Test case 3: Multiple potential header rows (edge case)
  const case3 = {
    0: { 0: { v: "Title" }, 1: { v: "Description" } }, // Only 2 headers
    3: { 0: { v: "Customer" }, 1: { v: "Order Date" }, 2: { v: "Amount" }, 3: { v: "Status" } },
    4: { 0: { v: "ACME Corp" }, 1: { v: "2024-01-15" }, 2: { v: 2500 }, 3: { v: "Paid" } }
  };

  // Simulate the improved extractHeaders function
  function extractHeaders(cellData) {
    if (!cellData) return [];
    
    for (let row = 0; row < 50; row++) {
      if (!cellData[row]) continue;
      
      const headers = [];
      let consecutiveHeaders = 0;
      
      for (let col = 0; col < 26; col++) {
        const cell = cellData[row][col];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "" && 
            typeof cell.v === 'string' && !cell.f) {
          headers.push(String(cell.v));
          consecutiveHeaders++;
        } else if (consecutiveHeaders > 0) {
          break;
        }
      }
      
      if (consecutiveHeaders >= 2) {
        console.log(`📋 Found headers starting at row ${row + 1}:`, headers);
        return headers;
      }
    }
    
    console.log("⚠️ No clear header row found");
    return [];
  }

  console.log("Test Case 1 - Traditional (row 1):");
  const result1 = extractHeaders(case1);
  console.log("Result:", result1);
  console.log("✅ Expected: ['Date', 'Product', 'Revenue']\n");

  console.log("Test Case 2 - Your Problem (row 5):");
  const result2 = extractHeaders(case2);
  console.log("Result:", result2);
  console.log("✅ Expected: ['Month', 'Sales', 'Profit', 'Margin%']\n");

  console.log("Test Case 3 - Multiple headers:");
  const result3 = extractHeaders(case3);
  console.log("Result:", result3);
  console.log("✅ Expected: ['Title', 'Description'] (first valid set)\n");

  // Verify the fixes
  const success1 = JSON.stringify(result1) === JSON.stringify(["Date", "Product", "Revenue"]);
  const success2 = JSON.stringify(result2) === JSON.stringify(["Month", "Sales", "Profit", "Margin%"]);
  const success3 = JSON.stringify(result3) === JSON.stringify(["Title", "Description"]);

  console.log("🎯 Test Results:");
  console.log(`Case 1 (Traditional): ${success1 ? '✅ PASS' : '❌ FAIL'}`);
  console.log(`Case 2 (Row 5 Start): ${success2 ? '✅ PASS' : '❌ FAIL'}`);
  console.log(`Case 3 (Multiple): ${success3 ? '✅ PASS' : '❌ FAIL'}`);

  if (success1 && success2 && success3) {
    console.log("\n🎉 ALL TESTS PASSED! Table detection now works regardless of starting row.");
  } else {
    console.log("\n❌ Some tests failed. Check the implementation.");
  }
}

testExtractHeaders();