function debugBreezeResponse() {
  const config = getConfig();
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/tags`;
  
  console.log('Testing URL:', url);
  console.log('Using API key:', config.BREEZE_API_KEY.substring(0, 8) + '...');
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('Response code:', responseCode);
    console.log('Response length:', responseText.length);
    console.log('First 200 characters:', responseText.substring(0, 200));
    
    if (responseText.length === 0) {
      console.log('❌ Empty response from Breeze API');
      return;
    }
    
    // Try to parse as JSON
    try {
      const jsonData = JSON.parse(responseText);
      console.log('✅ Valid JSON received');
      console.log('Number of tags:', jsonData.length);
      if (jsonData.length > 0) {
        console.log('First tag:', jsonData[0]);
      }
    } catch (parseError) {
      console.log('❌ JSON parse error:', parseError.message);
      console.log('Full response text:', responseText);
    }
    
  } catch (error) {
    console.log('❌ Request failed:', error.message);
  }
}
function testTagFolders() {
  const config = getConfig();
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/tag_folders`;
  
  console.log('Testing tag folders endpoint:', url);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('Tag folders - Response code:', responseCode);
    console.log('Tag folders - Response length:', responseText.length);
    console.log('First 200 characters:', responseText.substring(0, 200));
    
    if (responseText.length > 0) {
      try {
        const jsonData = JSON.parse(responseText);
        console.log('✅ Tag folders JSON valid');
        console.log('Number of folders:', jsonData.length);
        if (jsonData.length > 0) {
          console.log('First folder:', jsonData[0]);
        }
      } catch (parseError) {
        console.log('❌ Tag folders JSON parse error:', parseError.message);
      }
    }
    
  } catch (error) {
    console.log('Tag folders request failed:', error.message);
  }
}

function testPeopleEndpoint() {
  const config = getConfig();
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/people`;
  
  console.log('Testing people endpoint:', url);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('People - Response code:', responseCode);
    console.log('People - Response length:', responseText.length);
    console.log('First 100 characters:', responseText.substring(0, 100));
    
  } catch (error) {
    console.log('People request failed:', error.message);
  }
}
function testTagsWithLimit() {
  const config = getConfig();
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/tags?limit=5`;
  
  console.log('Testing tags with limit parameter:', url);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('Tags with limit - Response code:', responseCode);
    console.log('Tags with limit - Response length:', responseText.length);
    console.log('Response text:', responseText);
    
  } catch (error) {
    console.log('Tags with limit failed:', error.message);
  }
}
function testTagById() {
  const config = getConfig();
  const tagId = "4905834"; // Aaron White's tag ID
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/tags/${tagId}`;
  
  console.log('Testing tag by ID:', url);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('Tag by ID - Response code:', responseCode);
    console.log('Tag by ID - Response length:', responseText.length);
    console.log('Response text:', responseText);
    
  } catch (error) {
    console.log('Tag by ID failed:', error.message);
  }
}
function testPeopleWithTags() {
  const config = getConfig();
  // Let's try getting all people and see if tag information comes with them
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/people?details=1`;
  
  console.log('Testing people with details:', url);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('People with details - Response code:', responseCode);
    console.log('People with details - Response length:', responseText.length);
    
    // Parse and look for tag information in the first few people
    try {
      const people = JSON.parse(responseText);
      console.log('Number of people:', people.length);
      
      // Check first person's structure
      if (people.length > 0) {
        const firstPerson = people[0];
        console.log('First person keys:', Object.keys(firstPerson));
        console.log('First person sample:', JSON.stringify(firstPerson, null, 2).substring(0, 500));
      }
    } catch (e) {
      console.log('JSON parse error:', e.message);
    }
    
  } catch (error) {
    console.log('People with details failed:', error.message);
  }
}
function checkScottStringerRecord() {
  const config = getConfig();
  const personId = "29767824";
  const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/people/${personId}`;
  
  console.log('Getting Scott Stringer record with person ID:', personId);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Api-Key': config.BREEZE_API_KEY,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('Response code:', responseCode);
    console.log('Response length:', responseText.length);
    
    if (responseCode === 200 && responseText.length > 0) {
      const person = JSON.parse(responseText);
      
      console.log('=== SCOTT STRINGER FULL RECORD ===');
      console.log(JSON.stringify(person, null, 2));
      
      // Look specifically for tag-related data
      const fullRecord = JSON.stringify(person).toLowerCase();
      if (fullRecord.includes('4942192') || fullRecord.includes('gcg') || fullRecord.includes('tag')) {
        console.log('🏷️ Found potential tag references!');
      } else {
        console.log('❌ No obvious tag references found in record');
      }
    }
    
  } catch (error) {
    console.log('Failed to get Scott record:', error.message);
  }
}
function testScottGCGTag() {
  const config = getConfig();
  const tagId = "4942192"; // Your GCG tag
  
  // Try multiple approaches to your tag
  const endpoints = [
    `/api/tags/${tagId}`,
    `/api/tags/${tagId}/people`,
    `/api/people?tag_id=${tagId}`,
    `/api/people?tags=${tagId}`
  ];
  
  endpoints.forEach(endpoint => {
    const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com${endpoint}`;
    console.log(`\nTesting: ${url}`);
    
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Api-Key': config.BREEZE_API_KEY,
          'Content-Type': 'application/json'
        }
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log(`${endpoint} - Code: ${responseCode}, Length: ${responseText.length}`);
      if (responseText.length > 0 && responseText.length < 1000) {
        console.log(`Response: ${responseText}`);
      } else if (responseText.length > 0) {
        console.log(`First 200 chars: ${responseText.substring(0, 200)}`);
      }
      
    } catch (error) {
      console.log(`${endpoint} - Error: ${error.message}`);
    }
  });
}
function compareTaggedVsAllPeople() {
  const config = getConfig();
  
  console.log('Comparing people with tag vs all people...');
  
  try {
    // Get all people
    const allPeopleResponse = UrlFetchApp.fetch(`https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/people`, {
      headers: { 'Api-Key': config.BREEZE_API_KEY }
    });
    const allPeople = JSON.parse(allPeopleResponse.getContentText());
    
    // Get people with your tag
    const taggedPeopleResponse = UrlFetchApp.fetch(`https://${config.BREEZE_SUBDOMAIN}.breezechms.com/api/people?tag_id=4942192`, {
      headers: { 'Api-Key': config.BREEZE_API_KEY }
    });
    const taggedPeople = JSON.parse(taggedPeopleResponse.getContentText());
    
    console.log('All people count:', allPeople.length);
    console.log('Tagged people count:', taggedPeople.length);
    
    if (allPeople.length === taggedPeople.length) {
      console.log('❌ Same count - tag filtering is NOT working');
    } else {
      console.log('✅ Different counts - tag filtering IS working!');
      
      // Let's see who's in your GCG
      console.log('First few people in your GCG:');
      taggedPeople.slice(0, 5).forEach(person => {
        console.log(`- ${person.first_name} ${person.last_name}`);
      });
    }
    
  } catch (error) {
    console.log('Comparison failed:', error.message);
  }
}
function testBreezeExactFormat() {
  const config = getConfig();
  
  // Try the exact format from Breeze docs
  const tests = [
    // Method 1: Query string format
    `/api/people?filter_json={"tag_id":"4942192"}`,
    
    // Method 2: Try without quotes
    `/api/people?filter_json={"tag_id":4942192}`,
    
    // Method 3: Try with tag name instead of ID
    `/api/people?filter_json={"tag_name":"GCG: Gene Cone & Scott Stringer"}`,
    
    // Method 4: Different parameter name
    `/api/people?tag_filter=4942192`,
    
    // Method 5: POST request with filter
    // We'll try this as GET first
    `/api/people?search={"tag_id":"4942192"}`
  ];
  
  tests.forEach((endpoint, index) => {
    const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com${endpoint}`;
    console.log(`\nTest ${index + 1}: ${endpoint}`);
    
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Api-Key': config.BREEZE_API_KEY,
          'Content-Type': 'application/json'
        }
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        try {
          const people = JSON.parse(responseText);
          console.log(`✅ Success! ${people.length} people returned`);
          
          if (people.length > 0 && people.length < 100) {
            console.log('First few people:');
            people.slice(0, 3).forEach(person => {
              console.log(`- ${person.first_name} ${person.last_name}`);
            });
          }
        } catch (e) {
          console.log(`❌ JSON parse error: ${e.message}`);
        }
      } else {
        console.log(`❌ HTTP ${responseCode}: ${responseText.substring(0, 100)}`);
      }
      
    } catch (error) {
      console.log(`❌ Request failed: ${error.message}`);
    }
  });
}
function testActiveMembersTag() {
  const config = getConfig();
  const activeMembersTagId = "3623944";
  
  // Test if we can get just active members
  const tests = [
    `/api/people?tag_id=${activeMembersTagId}`,
    `/api/people?tags=${activeMembersTagId}`,
    `/api/people?tag_filter=${activeMembersTagId}`,
    `/api/people?filter_json={"tag_id":"${activeMembersTagId}"}`
  ];
  
  console.log('Testing Active Members tag filtering...');
  
  tests.forEach((endpoint, index) => {
    const url = `https://${config.BREEZE_SUBDOMAIN}.breezechms.com${endpoint}`;
    console.log(`\nTest ${index + 1}: ${endpoint}`);
    
    try {
      const response = UrlFetchApp.fetch(url, {
        headers: { 'Api-Key': config.BREEZE_API_KEY }
      });
      
      const people = JSON.parse(response.getContentText());
      console.log(`Returned ${people.length} people`);
      
      if (people.length === 617) {
        console.log('🎉 SUCCESS! Got exactly 617 active members!');
      } else if (people.length === 2515) {
        console.log('❌ Still getting all people - filtering not working');
      } else {
        console.log(`🤔 Got ${people.length} people - unexpected count`);
      }
      
    } catch (error) {
      console.log(`❌ Failed: ${error.message}`);
    }
  });
}
