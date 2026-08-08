/**
 * VERIFICATION SCRIPT FOR VOUCHER UNIQUENESS
 */
function verifyVoucherUniqueness() {
  const testToken = 'YOUR_TEST_TOKEN'; // Replace with a valid session token
  
  console.log('--- STARTING VOUCHER UNIQUENESS VERIFICATION ---');
  
  const testVoucher = {
    payee: 'TEST PAYEE',
    accountOrMail: 'VOU-TEST-999',
    particular: 'Test uniqueness',
    grossAmount: 1000,
    categories: 'General'
  };
  
  try {
    // 1. Create first voucher
    console.log('Test 1: Creating first voucher with VOU-TEST-999');
    const res1 = createVoucher(testToken, testVoucher);
    console.log('Result 1:', JSON.stringify(res1));
    
    if (res1.success) {
      console.log('✅ First voucher created.');
      
      // 2. Try to create second voucher with SAME number
      console.log('Test 2: Creating second voucher with SAME number VOU-TEST-999');
      const res2 = createVoucher(testToken, testVoucher);
      console.log('Result 2:', JSON.stringify(res2));
      
      if (!res2.success && res2.error.includes('already exists')) {
        console.log('✅ Test 2 Passed: Duplicate blocked successfully.');
      } else {
        console.log('❌ Test 2 Failed: Duplicate was not blocked or error message mismatch.');
      }
      
      // 3. Cleanup: Delete the first one if it's a real sheet
      // deleteVoucher(testToken, res1.rowIndex);
    } else {
      console.log('❌ Test 1 Failed: Could not create initial test voucher.', res1.error);
    }
    
  } catch (e) {
    console.error('Verification Error:', e.message);
  }
  
  console.log('--- VERIFICATION COMPLETE ---');
}
