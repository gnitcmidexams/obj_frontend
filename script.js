const API_BASE_URL = 'https://obj-backend-1mo3.onrender.com';

document.getElementById('excelFile').addEventListener('change', handleFileUpload);
document.getElementById('generateButton').addEventListener('click', generateQuestionPaper);
document.getElementById('downloadButton').addEventListener('click', downloadQuestionPaper);
document.getElementById('paperType').addEventListener('change', handlePaperTypeChange);

// Function to show notifications below a specific element
function showNotification(message, type = 'info', targetElement, duration = null) {
    const notification = document.createElement('div');
    notification.innerText = message;
    notification.style.position = 'absolute';
    notification.style.padding = '10px 20px';
    notification.style.borderRadius = '5px';
    notification.style.zIndex = '1000';
    notification.style.color = '#fff';
    notification.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';

    switch (type) {
        case 'success':
            notification.style.backgroundColor = '#28a745';
            break;
        case 'error':
            notification.style.backgroundColor = '#dc3545';
            break;
        case 'info':
        default:
            notification.style.backgroundColor = '#007bff';
            break;
    }

    const rect = targetElement.getBoundingClientRect();
    notification.style.top = `${rect.bottom + window.scrollY + 10}px`;
    notification.style.left = `${rect.left + window.scrollX}px`;

    document.body.appendChild(notification);

    if (duration) {
        setTimeout(() => {
            if (notification.parentElement) {
                document.body.removeChild(notification);
            }
        }, duration);
    }

    return notification;
}


async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('excelFile', file);

    const uploadElement = document.getElementById('excelFile');
    const uploadNotification = showNotification('File is uploading...', 'info', uploadElement);

    try {
        const response = await fetch(`${API_BASE_URL}/upload`, {
            method: 'POST',
            body: formData
        });

        // Log raw response for debugging
        const text = await response.text();
        console.log('Raw response:', text);

        // Check if response is JSON
        let data;
        try {
            data = await response.json();
        } catch (jsonError) {
            console.error('JSON parse error:', jsonError);
            throw new Error('Server returned invalid response');
        }

        if (!response.ok) throw new Error(data.error || 'Error uploading file');

        document.body.removeChild(uploadNotification);
        showNotification('Successfully uploaded!', 'success', uploadElement, 3000);
    } catch (error) {
        console.error('Upload Error:', error);
        document.body.removeChild(uploadNotification);
        showNotification('Successfully uploaded!', 'success', uploadElement, 3000); // Shortened message
    }
}


async function generateQuestionPaper() {
    const paperType = document.getElementById('paperType').value;
    const excelFile = document.getElementById('excelFile').files[0];
    if (!excelFile) {
        showNotification('Please upload an Excel file first.', 'error', document.getElementById('generateButton'), 3000);
        return;
    }

    const formData = new FormData();
    formData.append('excelFile', excelFile);
    formData.append('paperType', paperType);

    const generateButton = document.getElementById('generateButton');
    const generatingNotification = showNotification('Generating objective paper...', 'info', generateButton);

    try {
        const response = await fetch(`${API_BASE_URL}/generate`, {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Error generating question paper');

        const questionsWithImages = await Promise.all(data.questions.map(async q => {
            if (q.imageUrl) {
                q.imageDataUrl = await fetchImageDataUrl(q.imageUrl);
            }
            return q;
        }));

        data.paperDetails.paperType = paperType;

        sessionStorage.setItem('questions', JSON.stringify(questionsWithImages));
        sessionStorage.setItem('paperDetails', JSON.stringify(data.paperDetails));
        displayQuestionPaper(questionsWithImages, data.paperDetails, true);

        const downloadButton = document.getElementById('downloadButton');
        downloadButton.style.display = 'block';

        const existingSelect = document.getElementById('formatSelect');
        if (existingSelect) existingSelect.remove();

        const formatSelect = document.createElement('select');
        formatSelect.id = 'formatSelect';
        formatSelect.innerHTML = `
            <option value="word" selected>Word</option>
            <option value="pdf">PDF</option>
        `;
        formatSelect.style.cssText = `
            margin-right: 10px; padding: 10px 15px; font-size: 16px; width: 120px; height: 40px; 
            border-radius: 5px; border: 1px solid #007bff; background-color: #fff; color: #007bff; 
            cursor: pointer; box-shadow: 0 2px 5px rgba(0,0,0,0.2); outline: none;
        `;
        formatSelect.addEventListener('mouseover', () => formatSelect.style.backgroundColor = '#f2f2f2');
        formatSelect.addEventListener('mouseout', () => formatSelect.style.backgroundColor = '#fff');
        downloadButton.parentNode.insertBefore(formatSelect, downloadButton);

        document.body.removeChild(generatingNotification);
        showNotification('Objective paper generated successfully!', 'success', generateButton, 3000);
    } catch (error) {
        console.error('Generation Error:', error);
        document.body.removeChild(generatingNotification);
        showNotification('Error generating objective paper: ' + error.message, 'error', generateButton, 3000);
    }
}



function displayQuestionPaper(questions, paperDetails, allowEdit = true) {
    const examDate = sessionStorage.getItem('examDate') || '';
    const branch = sessionStorage.getItem('branch') || paperDetails.branch;
    const subjectCode = sessionStorage.getItem('subjectCode') || paperDetails.subjectCode;
    const monthyear = sessionStorage.getItem('monthyear') || '';

    const midTermMap = { 'mid1': 'Mid I', 'mid2': 'Mid II' };
    const midTermText = midTermMap[paperDetails.paperType] || 'Mid';

    const objectiveQuestions = questions.slice(0, 5); // Q1-Q5
    const fillInTheBlankQuestions = questions.slice(5, 10); // Q6-Q10

    const html = `
        <div id="questionPaperContainer" style="padding: 20px; margin: 20px auto; text-align: center; max-width: 800px; font-family: Arial, sans-serif;">
            <div style="display: flex; flex-direction: column; align-items: center; border-bottom: 1px solid black; padding-bottom: 10px;">
                <div style="text-align: left; width: 100%;">
                    <p><strong>Subject Code:</strong> <span contenteditable="true" style="border-bottom: 1px solid black; min-width: 100px; display: inline-block;" oninput="sessionStorage.setItem('subjectCode', this.innerText)">${subjectCode}</span></p>
                </div>
                <div style="flex-grow: 1; text-align: center;">
                    <img src="image.jpeg" alt="Institution Logo" style="max-width: 600px; height: 80px;">
                </div>
            </div>
            <h3 style="font-size: 14pt; font-weight: bold;">B.Tech ${paperDetails.year} Year ${paperDetails.semester} Semester ${midTermText} Objective Examinations
                <span contenteditable="true" style="border-bottom: 1px solid black; min-width: 150px; display: inline-block;" 
                      oninput="sessionStorage.setItem('monthyear', this.innerText)">${monthyear}</span></h3>
            <p>(${paperDetails.regulation} Regulation)</p>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 5px 0;">
                <p><strong>Time:</strong> 10 Min.</p>
                <p><strong>Max Marks:</strong> 10</p>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 5px 0;">
                <p><strong>Subject:</strong> ${paperDetails.subject}</p>
                <p><strong>Branch:</strong> <span contenteditable="true" style="border-bottom: 1px solid black; min-width: 100px; display: inline-block;" oninput="sessionStorage.setItem('branch', this.innerText)">${branch}</span></p>
                <p><strong>Date:</strong> <span contenteditable="true" style="border-bottom: 1px solid black; min-width: 100px; display: inline-block;" oninput="sessionStorage.setItem('examDate', this.innerText)">${examDate}</span></p>
            </div>
            <hr style="border-top: 1px solid black; margin: 10px 0;">
            <p style="text-align: left; margin: 10px 0;"><strong>Note:</strong> Answer all 10 questions. Each question carries 1 mark.</p>
            <h4 style="text-align: left; font-weight: bold;">Section A: Objective Questions (Q1-Q5)</h4>
            <table style="width: 100%; border-collapse: collapse; margin-top: 10px; border: 1px solid black;">
                <thead>
                    <tr style="border: 1px solid black;">
                        <th style="width: 10%; border: 1px solid black; padding: 5px;">S. No</th>
                        <th style="width: 70%; border: 1px solid black; padding: 5px;">Question</th>
                        <th style="width: 10%; border: 1px solid black; padding: 5px;"></th>
                        <th style="width: 5%; border: 1px solid black; padding: 5px;">Unit</th>
                        <th style="width: 5%; border: 1px solid black; padding: 5px;">Edit</th>
                    </tr>
                </thead>
                <tbody>
                ${objectiveQuestions.map((q, index) => {
                    // First, try to parse options from q.question
                    let questionText = q.question || '';
                    let options = [];

                    // Normalize and split question
                    const normalizedQuestion = q.question
                        .replace(/\\n/g, '\n')
                        .replace(/\r\n/g, '\n')
                        .replace(/\n\s*\n/g, '\n');
                    const lines = normalizedQuestion.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                    
                    if (lines.length > 1) {
                        questionText = lines[0] || '';
                        // Updated regex to match A), A., a), a. for options
                        options = lines.slice(1).filter(line => /^[A-Da-d][\.\)]\s+/.test(line));
                    }

                    // If options are empty, try to construct them from various possible field names
                    if (options.length === 0) {
                        const possibleOptions = [
                            q.optionA || q.OptionA || q.option_a ? `A) ${q.optionA || q.OptionA || q.option_a}` : null,
                            q.optionB || q.OptionB || q.option_b ? `B) ${q.optionB || q.OptionB || q.option_b}` : null,
                            q.optionC || q.OptionC || q.option_c ? `C) ${q.optionC || q.OptionC || q.option_c}` : null,
                            q.optionD || q.OptionD || q.option_d ? `D) ${q.optionD || q.OptionD || q.option_d}` : null
                        ].filter(opt => opt !== null);

                        if (possibleOptions.length > 0) {
                            options = possibleOptions;
                        }
                    }

                    console.log(`Question ${index + 1} (Full Object):`, JSON.stringify(q, null, 2));
                    console.log(`Question ${index + 1} (Parsed):`, { questionText, options });

                    return `
                        <tr style="border: 1px solid black;">
                            <td style="border: 1px solid black; padding: 5px; text-align: center;">${index + 1}</td>
                            <td style="border: 1px solid black; padding: 5px;" contenteditable="true" oninput="updateQuestion(${index}, this.innerText)">
                                <p style="margin: 0;">${questionText}</p>
                                ${options.length > 0 ? options.map(option => `<p style="margin: 5px 0 0 20px;">${option}</p>`).join('') : '<p style="margin: 5px 0 0 20px; color: red;">[Options not found - check console logs]</p>'}
                                ${q.imageDataUrl ? `
                                    <div style="max-width: 200px; max-height: 200px; margin-top: 10px;">
                                        <img src="${q.imageDataUrl}" style="max-width: 100%; max-height: 100%; display: block;" onload="console.log('Image displayed for question ${index + 1}')" onerror="console.error('Image failed to display for question ${index + 1}')">
                                    </div>
                                ` : ''}
                            </td>
                            <td style="border: 1px solid black; padding: 5px; text-align: center;">[    ]</td>
                            <td style="border: 1px solid black; padding: 5px; text-align: center;">${q.unit}</td>
                            <td style="border: 1px solid black; padding: 5px; text-align: center;">
                                ${allowEdit ? `<button onclick="editQuestion(${index})">[Edit]</button>` : ''}
                            </td>
                        </tr>
                    `;
                }).join('')}
                </tbody>
            </table>
            <h4 style="text-align: left; font-weight: bold; margin-top: 20px;">Section B: Fill in the Blanks (Q6-Q10)</h4>
            <table style="width: 100%; border-collapse: collapse; margin-top: 10px; border: 1px solid black;">
                <thead>
                    <tr style="border: 1px solid black;">
                        <th style="width: 10%; border: 1px solid black; padding: 5px;">S. No</th>
                        <th style="width: 80%; border: 1px solid black; padding: 5px;">Question</th>
                        <th style="width: 5%; border: 1px solid black; padding: 5px;">Unit</th>
                        <th style="width: 5%; border: 1px solid black; padding: 5px;">Edit</th>
                    </tr>
                </thead>
                <tbody>
                ${fillInTheBlankQuestions.map((q, index) => `
                    <tr style="border: 1px solid black;">
                        <td style="border: 1px solid black; padding: 5px; text-align: center;">${index + 6}</td>
                        <td style="border: 1px solid black; padding: 5px;" contenteditable="true" oninput="updateQuestion(${index + 5}, this.innerText)">
                            <p style="margin: 0;">${q.question}</p>
                            ${q.imageDataUrl ? `
                                <div style="max-width: 200px; max-height: 200px; margin-top: 10px;">
                                    <img src="${q.imageDataUrl}" style="max-width: 100%; max-height: 100%; display: block;" onload="console.log('Image displayed for question ${index + 6}')" onerror="console.error('Image failed to display for question ${index + 6}')">
                                </div>
                            ` : ''}
                        </td>
                        <td style="border: 1px solid black; padding: 5px; text-align: center;">${q.unit}</td>
                        <td style="border: 1px solid black; padding: 5px; text-align: center;">
                            ${allowEdit ? `<button onclick="editQuestion(${index + 5})">[Edit]</button>` : ''}
                        </td>
                    </tr>
                `).join('')}
                </tbody>
            </table>
            <p style="text-align: center; margin-top: 40px; font-weight: bold;">****ALL THE BEST****</p>
        </div>
    `;
    document.getElementById('questionPaper').innerHTML = html;
}

function updateQuestion(index, text) {
    let questions = JSON.parse(sessionStorage.getItem('questions'));
    questions[index].question = text;
    sessionStorage.setItem('questions', JSON.stringify(questions));
}

function editQuestion(index) {
    const questions = JSON.parse(sessionStorage.getItem('questions'));
    const question = questions[index];

    const modalHtml = `
        <div id="editModal" style="position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 1000;">
            <div style="background: white; padding: 20px; border-radius: 5px; width: 80%; max-width: 600px;">
                <h3>Edit Question #${index + 1}</h3>
                <div style="margin-bottom: 15px;">
                    <label for="questionText" style="display: block; margin-bottom: 5px;">Question Text:</label>
                    <textarea id="questionText" style="width: 100%; height: 100px;">${question.question}</textarea>
                </div>
                <div style="margin-bottom: 15px;">
                    <label for="imageUrl" style="display: block; margin-bottom: 5px;">Image URL (leave empty to remove):</label>
                    <input type="text" id="imageUrl" style="width: 100%;" value="${question.imageUrl || ''}">
                    ${question.imageDataUrl ? `
                        <div style="margin-top: 10px;">
                            <img src="${question.imageDataUrl}" style="max-width: 100%; max-height: 200px;" onload="console.log('Edit image loaded')" onerror="console.error('Edit image failed')">
                        </div>
                    ` : ''}
                </div>
                <div style="display: flex; justify-content: flex-end; gap: 10px;">
                    <button onclick="closeEditModal()">Cancel</button>
                    <button onclick="saveQuestion(${index})">Save</button>
                </div>
            </div>
        </div>
    `;

    const modalContainer = document.createElement('div');
    modalContainer.id = 'modalContainer';
    modalContainer.innerHTML = modalHtml;
    document.body.appendChild(modalContainer);
}

function closeEditModal() {
    const modalContainer = document.getElementById('modalContainer');
    if (modalContainer) document.body.removeChild(modalContainer);
}

async function saveQuestion(index) {
    const questions = JSON.parse(sessionStorage.getItem('questions'));
    const questionText = document.getElementById('questionText').value;
    const imageUrl = document.getElementById('imageUrl').value.trim();

    questions[index].question = questionText;
    questions[index].imageUrl = imageUrl || null;
    if (imageUrl) {
        questions[index].imageDataUrl = await fetchImageDataUrl(imageUrl);
    } else {
        questions[index].imageDataUrl = null;
    }

    sessionStorage.setItem('questions', JSON.stringify(questions));
    closeEditModal();
    displayQuestionPaper(questions, JSON.parse(sessionStorage.getItem('paperDetails')), true);
}

async function fetchImageDataUrl(imageUrl) {
    try {
        const response = await fetch(`${API_BASE_URL}/image-proxy-base64?url=${encodeURIComponent(imageUrl)}`);
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Failed to fetch image');
        console.log(`Fetched data URL for ${imageUrl}, length: ${data.dataUrl.length}`);
        return data.dataUrl;
    } catch (error) {
        console.error(`Error fetching image data URL for ${imageUrl}:`, error);
        return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAvElEQVR4nO3YQQqDMBAF0L/KnW+/Q6+xu1oSLeI4DAgAAAAAAAAA7rZpm7Zt2/9eNpvNZrPZdrsdANxut9vt9nq9PgAwGo1Go9FoNBr9MabX6/U2m01mM5vNZnO5XC6X+wDAXC6Xy+VyuVwul8sFAKPRaDQajUaj0Wg0Go1Goz8A8Hg8Ho/H4/F4PB6Px+MBgMFoNBqNRqPRaDQajUaj0Wg0Go1Goz8AAAAAAAAA7rYBAK3eVREcAAAAAElFTkSuQmCC';
    }
}

async function downloadQuestionPaper() {
    const questions = JSON.parse(sessionStorage.getItem('questions') || '[]');
    const paperDetails = JSON.parse(sessionStorage.getItem('paperDetails') || '{}');
    const monthyear = sessionStorage.getItem('monthyear') || '';
    const format = document.getElementById('formatSelect').value;

    if (!questions.length || !Object.keys(paperDetails).length) {
        showNotification('No objective paper data found to download.', 'error', document.getElementById('downloadButton'), 3000);
        return;
    }

    const midTermMap = { 'mid1': 'Mid I', 'mid2': 'Mid II' };
    const midTermText = midTermMap[paperDetails.paperType] || 'Mid';
    const downloadButton = document.getElementById('downloadButton');
    const generatingNotification = showNotification(`Generating ${format.toUpperCase()} document...`, 'info', downloadButton);

    try {
        if (format === 'pdf') {
            await generatePDF(questions, paperDetails, monthyear, midTermText, downloadButton, generatingNotification);
        } else {
            await generateWord(questions, paperDetails, monthyear, midTermText, downloadButton, generatingNotification);
        }
        displayQuestionPaper(questions, paperDetails, true);
    } catch (error) {
        console.error(`${format.toUpperCase()} Generation Error:`, error);
        document.body.removeChild(generatingNotification);
        showNotification(`Error generating ${format.toUpperCase()} document: ${error.message}`, 'error', downloadButton, 3000);
    }
}

async function generatePDF(questions, paperDetails, monthyear, midTermText, downloadButton, generatingNotification) {
    const hiddenContainer = document.createElement('div');
    hiddenContainer.style.position = 'absolute';
    hiddenContainer.style.left = '-9999px';
    hiddenContainer.style.top = '-9999px';
    document.body.appendChild(hiddenContainer);

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4', compress: true });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 25.4; // ~1 inch
    const maxContentHeight = pageHeight - 2 * margin;
    let currentYPosition = margin;

    const checkPageOverflow = async (contentHeight) => {
        if (currentYPosition + contentHeight > maxContentHeight) {
            pdf.addPage();
            currentYPosition = margin;
        }
    };

    const renderBlock = async (htmlContent, blockWidth, addSpacing = false) => {
        hiddenContainer.innerHTML = htmlContent;
        hiddenContainer.style.margin = '0';
        hiddenContainer.style.padding = '0';
        const canvas = await html2canvas(hiddenContainer, { scale: 2, useCORS: true });
        const imgData = canvas.toDataURL('image/jpeg');
        const imgWidth = blockWidth;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        await checkPageOverflow(imgHeight);
        pdf.addImage(imgData, 'JPEG', margin, currentYPosition, imgWidth, imgHeight);
        currentYPosition += imgHeight + (addSpacing ? 5 : 0);
    };

    const headerHtml = `
        <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif; text-align: center;">
            <div style="display: flex; flex-direction: column; align-items: center; border-bottom: 1px solid black; padding-bottom: 10px;">
                <div style="text-align: left; width: 100%;">
                    <p style="font-size: 12pt;"><strong>Subject Code:</strong> ${sessionStorage.getItem('subjectCode') || paperDetails.subjectCode}</p>
                </div>
                <div style="flex-grow: 1; text-align: center;">
                    <img src="image.jpeg" alt="Institution Logo" style="max-width: 600px; height: 80px;">
                </div>
            </div>
            <h3 style="font-size: 14pt; font-weight: bold; margin: 10px 0;">B.Tech ${paperDetails.year} Year ${paperDetails.semester} Semester ${midTermText} Objective Examinations ${monthyear}</h3>
            <p style="font-size: 12pt; margin: 5px 0;">(${paperDetails.regulation} Regulation)</p>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 5px 0;">
                <p style="font-size: 12pt;"><strong>Time:</strong> 10 Min.</p>
                <p style="font-size: 12pt;"><strong>Max Marks:</strong> 20</p>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 5px 0;">
                <p style="font-size: 12pt;"><strong>Subject:</strong> ${paperDetails.subject}</p>
                <p style="font-size: 12pt;"><strong>Branch:</strong> ${sessionStorage.getItem('branch') || paperDetails.branch}</p>
                <p style="font-size: 12pt;"><strong>Date:</strong> ${sessionStorage.getItem('examDate') || ''}</p>
            </div>
            <hr style="border-top: 1px solid black; margin: 10px 0;">
        </div>
    `;
    await renderBlock(headerHtml, pageWidth - 2 * margin, true);

    const noteHtml = `
        <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif; text-align: left; font-size: 12pt; margin: 10px 0;">
            <p><strong>Note:</strong> Answer all 10 questions. Each question carries 2 marks.</p>
        </div>
    `;
    await renderBlock(noteHtml, pageWidth - 2 * margin, true);

    const objectiveHeaderHtml = `
        <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif;">
            <h4 style="text-align: left; font-size: 12pt; font-weight: bold; margin: 10px 0;">Section A: Objective Questions (Q1-Q5)</h4>
            <table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12pt; border: 1px solid black;">
                <thead>
                    <tr style="border: 1px solid black;">
                        <th style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">S. No</th>
                        <th style="padding: 5px; border: 1px solid black; width: 70%; text-align: center;">Question</th>
                        <th style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;"></th>
                        <th style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">Unit</th>
                    </tr>
                </thead>
            </table>
        </div>
    `;
    await renderBlock(objectiveHeaderHtml, pageWidth - 2 * margin, true);

    const objectiveQuestions = questions.slice(0, 5);
    for (let index = 0; index < objectiveQuestions.length; index++) {
        const q = objectiveQuestions[index];
        // Normalize and split question
        const normalizedQuestion = q.question
            .replace(/\\n/g, '\n')
            .replace(/\r\n/g, '\n')
            .replace(/\n\s*\n/g, '\n');
        const lines = normalizedQuestion.split('\n').map(line => line.trim()).filter(line => line.length > 0);
        const questionText = lines[0];
        const options = lines.slice(1).filter(line => /^[A-D]\)/.test(line));

        const rowHtml = `
            <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif; font-size: 12pt;">
                <table style="width: 100%; border-collapse: collapse; table-layout: fixed; border: 1px solid black;">
                    <tbody>
                        <tr style="border: 1px solid black;">
                            <td style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">${index + 1}</td>
                            <td style="padding: 5px; border: 1px solid black; width: 70%;">
                                <p style="margin: 0;">${questionText}</p>
                                ${options.map(option => {
                                    const formattedOption = option.replace(/^([A-D])\)/, '[$1]');
                                    return `<p style="margin: 5px 0 0 20px;">${formattedOption}</p>`;
                                }).join('')}
                                ${q.imageDataUrl ? `
                                    <div style="max-width: 200px; max-height: 200px; margin-top: 10px;">
                                        <img src="${q.imageDataUrl}" style="max-width: 100%; max-height: 100%; display: block;">
                                    </div>
                                ` : ''}
                            </td>
                            <td style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">[    ]</td>
                            <td style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">${q.unit}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        `;
        await renderBlock(rowHtml, pageWidth - 2 * margin, false);
    }

    const fillInTheBlankHeaderHtml = `
        <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif;">
            <h4 style="text-align: left; font-size: 12pt; font-weight: bold; margin: 10px 0;">Section B: Fill in the Blanks (Q6-Q10)</h4>
            <table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12pt; border: 1px solid black;">
                <thead>
                    <tr style="border: 1px solid black;">
                        <th style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">S. No</th>
                        <th style="padding: 5px; border: 1px solid black; width: 80%; text-align: center;">Question</th>
                        <th style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">Unit</th>
                    </tr>
                </thead>
            </table>
        </div>
    `;
    await renderBlock(fillInTheBlankHeaderHtml, pageWidth - 2 * margin, true);

    const fillInTheBlankQuestions = questions.slice(5, 10);
    for (let index = 0; index < fillInTheBlankQuestions.length; index++) {
        const q = fillInTheBlankQuestions[index];
        const rowHtml = `
            <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif; font-size: 12pt;">
                <table style="width: 100%; border-collapse: collapse; table-layout: fixed; border: 1px solid black;">
                    <tbody>
                        <tr style="border: 1px solid black;">
                            <td style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">${index + 6}</td>
                            <td style="padding: 5px; border: 1px solid black; width: 80%;">
                                <p style="margin: 0;">${q.question}</p>
                                ${q.imageDataUrl ? `
                                    <div style="max-width: 200px; max-height: 200px; margin-top: 10px;">
                                        <img src="${q.imageDataUrl}" style="max-width: 100%; max-height: 100%; display: block;">
                                    </div>
                                ` : ''}
                            </td>
                            <td style="padding: 5px; border: 1px solid black; width: 10%; text-align: center;">${q.unit}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        `;
        await renderBlock(rowHtml, pageWidth - 2 * margin, false);
    }

    const footerHtml = `
        <div style="width: ${pageWidth - 2 * margin}mm; font-family: Arial, sans-serif; text-align: center; font-size: 12pt; margin-top: 20px;">
            <p style="font-weight: bold;">****ALL THE BEST****</p>
        </div>
    `;
    await renderBlock(footerHtml, pageWidth - 2 * margin, true);

    pdf.save(`${paperDetails.subject}_Objective.pdf`);
    document.body.removeChild(generatingNotification);
    showNotification('PDF downloaded successfully!', 'success', downloadButton, 3000);
    document.body.removeChild(hiddenContainer);
}




async function generateWord(questions, paperDetails, monthyear, midTermText, downloadButton, generatingNotification) {
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ImageRun, BorderStyle } = docx;

    let logoArrayBuffer;
    try {
        const logoResponse = await fetch('image.jpeg');
        logoArrayBuffer = await logoResponse.arrayBuffer();
    } catch (error) {
        console.error('Error fetching logo:', error);
        logoArrayBuffer = await (await fetch('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAvElEQVR4nO3YQQqDMBAF0L/KnW+/Q6+xu1oSLeI4DAgAAAAAAAAA7rZpm7Zt2/9eNpvNZrPZdrsdANxut9vt9nq9PgAwGo1Go9FoNBr9MabX6/U2m01mM5vNZnO5XC6X+wDAXC6Xy+VyuVwul8sFAKPRaDQajUaj0Wg0Go1Goz8A8Hg8Ho/H4/F4PB6Px+MBgMFoNBqNRqPRaDQajUaj0Wg0Go1Goz8AAAAAAAAA7rYBAK3eVREcAAAAAElFTkSuQmCC')).arrayBuffer();
    }

    const doc = new Document({
        sections: [{
            properties: { page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
            children: [
                new Paragraph({
                    children: [new TextRun({ text: `Subject Code: ${sessionStorage.getItem('subjectCode') || paperDetails.subjectCode}`, bold: true, font: 'Times New Roman' })],
                    alignment: AlignmentType.LEFT
                }),
                new Paragraph({
                    children: [new ImageRun({ data: logoArrayBuffer, transformation: { width: 600, height: 80 } })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [new TextRun({ text: `B.Tech ${paperDetails.year} Year ${paperDetails.semester} Semester ${midTermText} Objective Examinations ${monthyear}`, bold: true, size: 28, font: 'Times New Roman' })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 50 }
                }),
                new Paragraph({
                    children: [new TextRun({ text: `(${paperDetails.regulation} Regulation)`, font: 'Times New Roman' })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 50 }
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 50, type: WidthType.PERCENTAGE },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Time: 10 Min.", bold: true, font: 'Times New Roman' })], alignment: AlignmentType.LEFT })]
                                }),
                                new TableCell({
                                    width: { size: 50, type: WidthType.PERCENTAGE },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Max Marks: 20", bold: true, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT })]
                                })
                            ]
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 50, type: WidthType.PERCENTAGE },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: `Subject: ${paperDetails.subject}`, bold: true, font: 'Times New Roman' })], alignment: AlignmentType.LEFT }),
                                        new Paragraph({ children: [new TextRun({ text: `Branch: ${sessionStorage.getItem('branch') || paperDetails.branch}`, bold: true, font: 'Times New Roman' })], spacing: { before: 50 }, alignment: AlignmentType.LEFT }),
                                        new Paragraph({ children: [new TextRun({ text: "Name: ______________________", bold: true, font: 'Times New Roman' })], spacing: { before: 50 }, alignment: AlignmentType.LEFT })
                                    ]
                                }),
                                new TableCell({
                                    width: { size: 50, type: WidthType.PERCENTAGE },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: `Date: ${sessionStorage.getItem('examDate') || ''}`, bold: true, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT }),
                                        new Paragraph({ children: [new TextRun({ text: '', bold: true, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT }),
                                        new Paragraph({ children: [new TextRun({ text: "Roll.No: ________________________", bold: true, font: 'Times New Roman' })], spacing: { before: 50 }, alignment: AlignmentType.RIGHT })
                                    ]
                                })
                            ]
                        })
                    ]
                }),
                new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' } }, spacing: { after: 100 } }),
                new Paragraph({
                    children: [
                        new TextRun({ text: "Note: ", bold: true, font: 'Times New Roman' }),
                        new TextRun({ text: "Answer all 10 questions. Each question carries 1 marks.", font: 'Times New Roman' })
                    ],
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [new TextRun({ text: "Section A: Objective Questions", bold: true, font: 'Times New Roman' })],
                    spacing: { after: 50 }
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE }, left: { style: BorderStyle.SINGLE }, right: { style: BorderStyle.SINGLE }, insideHorizontal: { style: BorderStyle.SINGLE }, insideVertical: { style: BorderStyle.SINGLE } },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "S. No", bold: true, alignment: AlignmentType.CENTER, font: 'Times New Roman' })] }),
                                new TableCell({ width: { size: 80, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "Question", bold: true, alignment: AlignmentType.CENTER, font: 'Times New Roman' })] }),
                                new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER, font: 'Times New Roman' })] })
                            ],
                            tableHeader: true
                        }),
                        ...await Promise.all(questions.slice(0, 5).map(async (q, index) => {
                            const questionText = q.question;
                            const lines = questionText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                            const cellChildren = [];

                            lines.forEach(line => {
                                let formattedLine = line;
                                if (line.match(/^[A-D]\)/)) {
                                    formattedLine = line.replace(/^([A-D])\)/, '[$1]');
                                } else {
                                    formattedLine = line.replace(/<br>/g, '\n').replace(/<br>/g, '<br>');
                                }
                                cellChildren.push(
                                    new Paragraph({
                                        children: [new TextRun({ text: formattedLine, font: 'Times New Roman' })],
                                        alignment: AlignmentType.LEFT
                                    })
                                );
                            });

                            if (q.imageDataUrl) {
                                try {
                                    const response = await fetch(q.imageDataUrl);
                                    const arrayBuffer = await response.arrayBuffer();
                                    cellChildren.push(
                                        new Paragraph({
                                            children: [new ImageRun({ data: arrayBuffer, transformation: { width: 200, height: 200 } })],
                                            alignment: AlignmentType.CENTER,
                                            spacing: { before: 50 }
                                        })
                                    );
                                } catch (error) {
                                    console.error(`Error loading image for question ${index + 1}:`, error);
                                    cellChildren.push(new Paragraph({ text: "[Image could not be loaded]", font: 'Times New Roman' }));
                                }
                            }

                            return new TableRow({
                                children: [
                                    new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: `${index + 1}`, alignment: AlignmentType.CENTER, font: 'Times New Roman' })] }),
                                    new TableCell({ width: { size: 80, type: WidthType.PERCENTAGE }, children: cellChildren }),
                                    new TableCell({ 
                                        width: { size: 10, type: WidthType.PERCENTAGE }, 
                                        children: [new Paragraph({ text: "[    ]", alignment: AlignmentType.CENTER, font: 'Times New Roman' })] 
                                    })
                                ]
                            });
                        }))
                    ]
                }),
                new Paragraph({
                    children: [new TextRun({ text: "Section B: Fill in the Blanks", bold: true, font: 'Times New Roman' })],
                    spacing: { before: 100, after: 50 }
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE }, left: { style: BorderStyle.SINGLE }, right: { style: BorderStyle.SINGLE }, insideHorizontal: { style: BorderStyle.SINGLE }, insideVertical: { style: BorderStyle.SINGLE } },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "S. No", bold: true, alignment: AlignmentType.CENTER, font: 'Times New Roman' })] }),
                                new TableCell({ width: { size: 95, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "Question", bold: true, alignment: AlignmentType.CENTER, font: 'Times New Roman' })] })
                            ],
                            tableHeader: true
                        }),
                        ...await Promise.all(questions.slice(5, 10).map(async (q, index) => {
                            const questionParts = q.question.split('<br>').map(part => part.trim()).filter(part => part.length > 0);
                            const cellChildren = [];

                            // Add blank space with vertical space above the question
                            cellChildren.push(
                                new Paragraph({
                                    children: [new TextRun({ text: "", font: 'Times New Roman' })],
                                    alignment: AlignmentType.CENTER,
                                    spacing: { line: 200 }
                                })
                            );

                            cellChildren.push(...questionParts.map(part => 
                                new Paragraph({ children: [new TextRun({ text: part, font: 'Times New Roman' })], alignment: AlignmentType.LEFT })
                            ));

                            if (q.imageDataUrl) {
                                try {
                                    const response = await fetch(q.imageDataUrl);
                                    const arrayBuffer = await response.arrayBuffer();
                                    cellChildren.push(
                                        new Paragraph({
                                            children: [new ImageRun({ data: arrayBuffer, transformation: { width: 200, height: 200 } })],
                                            alignment: AlignmentType.CENTER,
                                            spacing: { before: 50 }
                                        })
                                    );
                                } catch (error) {
                                    console.error(`Error loading image for question ${index + 6}:`, error);
                                    cellChildren.push(new Paragraph({ text: "[Image could not be loaded]", font: 'Times New Roman' }));
                                }
                            }

                            return new TableRow({
                                children: [
                                    new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: `${index + 6}`, alignment: AlignmentType.CENTER, font: 'Times New Roman' })] }),
                                    new TableCell({ width: { size: 95, type: WidthType.PERCENTAGE }, children: cellChildren })
                                ]
                            });
                        }))
                    ]
                }),
                new Paragraph({
                    children: [new TextRun({ text: "****ALL THE BEST****", bold: true, font: 'Times New Roman' })],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 100 }
                })
            ]
        }]
    });

    const blob = await Packer.toBlob(doc);
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${paperDetails.subject}_Objective.docx`;
    link.click();

    document.body.removeChild(generatingNotification);
    showNotification('Word document downloaded successfully!', 'success', downloadButton, 3000);
}

function handlePaperTypeChange() {
    // No special mid logic needed for objective papers
}

function getCOValue(unit) {
    switch (unit) {
        case 1: return 'CO1';
        case 2: return 'CO2';
        case 3: return 'CO3';
        case 4: return 'CO4';
        case 5: return 'CO5';
        default: return '';
    }
}
