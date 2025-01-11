const xlsx = require("xlsx");
const { GoogleGenerativeAI } = require("@google/generative-ai");


async function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms)); // Sleep for a given time in milliseconds
}

function apikeys(count){
    keys = [
    ]
    return keys[count%10]
}

async function classifyProduct(productName, categoryMap,apikey) {
    const genAI = new GoogleGenerativeAI(apikey);
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" })

    try {
        // Step 1: Identify the main category
        const mainCategories = Object.keys(categoryMap);
        const prompt1 = `
        Product: "${productName}"
        Categories: ${JSON.stringify(mainCategories)}
        Note: The product names are primarily Indian. If you cannot confidently determine the correct main category, return an empty string.
        Be 100% sure before assigning a main category.
        Return only the main category as a single line of text.
        `;

        const result1 = await model.generateContent(prompt1);
        const mainCategory = result1.response.text().trim();

        // Check if the main category is valid
        if (!mainCategory || !mainCategories.includes(mainCategory)) {
            return { mainCategory: "", subcategory: "" };
        }

        // Step 2: Identify the subcategory
        const subcategories = categoryMap[mainCategory] || [];
        const prompt2 = `
        Product: "${productName}"
        Main Category: "${mainCategory}"
        Subcategories: ${JSON.stringify(subcategories)}
        Note: The product names and subcategories are primarily Indian. If you cannot confidently determine the correct subcategory, return an empty string.
        Be 100% sure before assigning a subcategory.
        Return only the subcategory as a single line of text.
        `;
        // sleep(2000)
        const result2 = await model.generateContent(prompt2);
        const subcategory = result2.response.text().trim();

        // Check if the subcategory is valid
        if (!subcategory || !subcategories.includes(subcategory)) {
            return { mainCategory, subcategory: "" };
        }

        return { mainCategory, subcategory };
    } catch (error) {
        console.error(`Error classifying product: ${productName}`, error);
        return { mainCategory: "", subcategory: "" }; // Return empty values on error
    }
}


async function processExcel(filePath, outputPath, categoryMap) {

    // Read the Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert to 2D array
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    let counter = 0;
    // Process each row, skipping the header
    for (let i = 2497; i < data.length; i++) {
        
        const row = data[i];
        const productName = row[1]; // Product name in 0th column
        const subcategory = row[12]; // Subcategory in 12th column
        const mainCategory = row[13]; // Main category in 13th column

        if (mainCategory) continue; // Skip rows where the main category is already present

        counter++;

        const apikey = apikeys(counter);

        console.log(`Processing row ${i + 1}-${counter} : ${productName} ... apikey: ${apikey}`);
        
        const { mainCategory: newMainCategory, subcategory: newSubcategory } = 
            await classifyProduct(productName, categoryMap,apikey);

        // Update the row with AI results
        row[12] = newSubcategory || subcategory;
        row[13] = newMainCategory || mainCategory;
        
        if(counter == 500)break
        await sleep(1000);
    }

    // Write the updated data back to a new Excel file
    const newSheet = xlsx.utils.aoa_to_sheet(data);
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
    xlsx.writeFile(newWorkbook, outputPath);

    console.log(`File saved to ${outputPath}`);
}

// Define your category map here
const handwritten_categories = {
    "Food & Beverages": ["Packaged Food Item", "Snacks", "Soft Drink", "Juices", "Dairy", "Ghee and Oil", "Dal & Pulses", "Dry Fruits", "Frozen Food", "Beverages"],
    "Personal Care": ["Toiletries", "Skincare", "Haircare", "Oral Care", "Body Care"],
    "Household": ["Laundry Detergents", "Cleaning Agents", "Air Fresheners", "Disinfectants"],
    "Health & Wellness": [],
    "Baby and Childcare": ["Diapers", "Baby Food", "Baby Wipes", "Baby Skin Products"],
    "Pet Care": ["Pet Food", "Pet Grooming Products", "Pet Accessories"],
    "Confectionery and Chocolate": ["Chocolate", "Candies and Sweets", "Gum and Chewing"],
    "Tobacco and Cigarettes": [],
    "Grocery": ["Masala & Spices"]
};

// Run the process
processExcel("demoproduct.xls", "demoproduct-updated_file.xls", handwritten_categories);

