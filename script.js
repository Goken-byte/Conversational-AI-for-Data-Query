class ExcelAIQuery {
    constructor() {
        this.excelData = null;
        this.apiKey = '';
        this.isProcessing = false;
        this.headers = [];
        this.currentLoadingMessage = null;
        this.isChatMinimized = false;
        
        this.initializeEventListeners();
        this.loadSavedApiKey();
    }

    initializeEventListeners() {
        // File upload
        document.getElementById('excelFile').addEventListener('change', (e) => {
            this.handleFileUpload(e.target.files[0]);
        });

        // API key input
        document.getElementById('apiKey').addEventListener('input', (e) => {
            this.apiKey = e.target.value.trim();
            this.updateUIState();
        });

        // Save API key button
        document.getElementById('saveApiKey').addEventListener('click', () => {
            this.saveApiKey();
        });

        // Send button and Enter key
        document.getElementById('sendBtn').addEventListener('click', () => {
            this.sendMessage();
        });

        document.getElementById('userInput').addEventListener('keypress', (e) => {
            if (e.key === 'Enter' && !this.isProcessing) {
                this.sendMessage();
            }
        });

        // Minimize chat
        document.getElementById('minimizeBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggleChatMinimize();
        });

        // Click header to maximize
        document.querySelector('.chat-header').addEventListener('click', () => {
            if (this.isChatMinimized) {
                this.toggleChatMinimize();
            }
        });
    }

    loadSavedApiKey() {
        const savedApiKey = localStorage.getItem('openrouterApiKey');
        if (savedApiKey) {
            this.apiKey = savedApiKey;
            document.getElementById('apiKey').value = savedApiKey;
            this.updateUIState();
        }
    }

    saveApiKey() {
        if (this.apiKey) {
            localStorage.setItem('openrouterApiKey', this.apiKey);
            this.addMessage('System: API key saved successfully!', 'bot');
        } else {
            this.addMessage('System: Please enter an API key first.', 'bot');
        }
    }

    toggleChatMinimize() {
        const chatContainer = document.querySelector('.chat-floating-container');
        this.isChatMinimized = !this.isChatMinimized;
        
        if (this.isChatMinimized) {
            chatContainer.classList.add('minimized');
            document.getElementById('minimizeBtn').textContent = '+';
        } else {
            chatContainer.classList.remove('minimized');
            document.getElementById('minimizeBtn').textContent = '−';
            // Scroll to bottom when maximizing
            const chatMessages = document.getElementById('chatMessages');
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
    }

    handleFileUpload(file) {
        if (!file) return;

        document.getElementById('fileName').textContent = `Selected: ${file.name}`;
        this.addMessage(`System: Reading ${file.name}...`, 'bot');
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Get first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON with headers
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                this.excelData = jsonData;
                
                // Extract headers
                if (jsonData.length > 0) {
                    this.headers = Object.keys(jsonData[0]);
                }
                
                this.addMessage(`System: Excel file loaded successfully! Found ${jsonData.length} rows with columns: ${this.headers.join(', ')}`, 'bot');
                this.updateTableInfo(`Loaded ${jsonData.length} rows with ${this.headers.length} columns`);
                this.displayDataTable(this.excelData, `Complete Dataset (${this.excelData.length} rows)`);
                this.updateUIState();
                
            } catch (error) {
                this.addMessage(`Error: Failed to read Excel file - ${error.message}`, 'bot');
                console.error('File reading error:', error);
            }
        };
        
        reader.onerror = () => {
            this.addMessage('Error: Failed to read the file. Please try again.', 'bot');
        };
        
        reader.readAsArrayBuffer(file);
    }

    updateTableInfo(message) {
        document.getElementById('tableInfo').textContent = message;
    }

    updateUIState() {
        const hasData = this.excelData && this.excelData.length > 0;
        const hasApiKey = this.apiKey.length > 0;
        const isReady = hasData && hasApiKey && !this.isProcessing;

        document.getElementById('userInput').disabled = !isReady;
        document.getElementById('sendBtn').disabled = !isReady;
        
        if (!hasData) {
            document.getElementById('userInput').placeholder = 'Upload Excel file first...';
        } else if (!hasApiKey) {
            document.getElementById('userInput').placeholder = 'Enter API key...';
        } else {
            document.getElementById('userInput').placeholder = 'Ask about your data...';
        }
    }

    addMessage(text, sender) {
        const chatMessages = document.getElementById('chatMessages');
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${sender}-message`;
        messageDiv.textContent = text;
        chatMessages.appendChild(messageDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        return messageDiv;
    }

    removeMessage(messageElement) {
        if (messageElement && messageElement.parentNode) {
            messageElement.parentNode.removeChild(messageElement);
        }
    }

    async sendMessage() {
        const userInput = document.getElementById('userInput');
        const message = userInput.value.trim();

        if (!message || this.isProcessing) return;

        // Add user message
        this.addMessage(message, 'user');
        userInput.value = '';
        this.isProcessing = true;
        this.updateUIState();

        // Add loading indicator
        this.currentLoadingMessage = this.addMessage('AI: Analyzing your data...', 'bot');

        try {
            const timeoutPromise = new Promise((_, reject) => {
                setTimeout(() => reject(new Error('Request timeout after 30 seconds')), 30000);
            });

            const queryPromise = this.queryOpenRouter(message);
            const response = await Promise.race([queryPromise, timeoutPromise]);
            
            this.removeMessage(this.currentLoadingMessage);
            this.processAIResponse(response, message);
            
        } catch (error) {
            this.removeMessage(this.currentLoadingMessage);
            this.addMessage(`Error: ${error.message}`, 'bot');
            console.error('AI Query error:', error);
        } finally {
            this.isProcessing = false;
            this.currentLoadingMessage = null;
            this.updateUIState();
        }
    }

    async queryOpenRouter(query) {
        if (!this.excelData || this.excelData.length === 0) {
            throw new Error('No Excel data available');
        }

        if (!this.apiKey) {
            throw new Error('OpenRouter API key is required');
        }

        const sampleData = this.excelData.slice(0, 3);

        const prompt = `You are a data analysis assistant. Analyze this Excel data and respond in JSON format.

DATA STRUCTURE:
- Headers: ${JSON.stringify(this.headers)}
- Total Records: ${this.excelData.length}
- Sample Data: ${JSON.stringify(sampleData)}

USER QUERY: "${query}"

RESPONSE FORMAT (JSON only):
{
    "answer": "Clear explanation of findings",
    "data_subset": [array of relevant data rows],
    "summary": "Brief summary of results",
    "filter_used": "Description of how data was filtered"
}

INSTRUCTIONS:
1. Analyze the query and filter the data accordingly
2. Include the actual data rows in "data_subset"
3. If showing filtered data, include ALL columns for each row
4. If no specific filter, return meaningful subset (max 20 rows)
5. For calculations, include the calculation results
6. Always return some data in "data_subset"

Respond with ONLY the JSON object, no additional text.`;

        const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${this.apiKey}`,
                'HTTP-Referer': window.location.href,
                'X-Title': 'Excel AI Query'
            },
            body: JSON.stringify({
                model: 'anthropic/claude-3-haiku', // Changed to Claude 3 Haiku
                messages: [{
                    role: 'user',
                    content: prompt
                }],
                max_tokens: 2000,
                temperature: 0.1
            })
        });

        if (!response.ok) {
            let errorMessage = `API Error: ${response.status}`;
            
            try {
                const errorData = await response.json();
                if (errorData.error?.message) {
                    errorMessage = errorData.error.message;
                }
                
                if (response.status === 401) {
                    errorMessage = 'Invalid API key';
                } else if (response.status === 402) {
                    errorMessage = 'Please add credits to your OpenRouter account';
                } else if (response.status === 429) {
                    errorMessage = 'Rate limit exceeded';
                }
            } catch (e) {
                errorMessage = `API Error: ${response.status} ${response.statusText}`;
            }
            
            throw new Error(errorMessage);
        }

        const data = await response.json();
        
        if (!data.choices || !data.choices[0] || !data.choices[0].message) {
            throw new Error('Invalid response format from AI service');
        }

        return data.choices[0].message.content;
    }

    processAIResponse(aiResponse, originalQuery) {
        console.log('Raw AI Response:', aiResponse);
        
        try {
            // Clean the response and parse JSON
            const cleanResponse = aiResponse.replace(/```json\s?/g, '').replace(/```\s?/g, '').trim();
            const result = JSON.parse(cleanResponse);
            
            // Display the answer in chat
            this.addMessage(`AI: ${result.answer}`, 'bot');
            
            // Display summary if available
            if (result.summary) {
                this.addMessage(`AI: Summary: ${result.summary}`, 'bot');
            }
            
            // Always update the table with the data subset
            if (result.data_subset && result.data_subset.length > 0) {
                const filterInfo = result.filter_used || `Filtered results for: "${originalQuery}"`;
                this.displayDataTable(result.data_subset, `${filterInfo} (${result.data_subset.length} rows)`);
            } else {
                // If no specific data subset, show smart filtered data based on query
                const smartData = this.getSmartFilteredData(originalQuery);
                this.displayDataTable(smartData, `Results for: "${originalQuery}" (${smartData.length} rows)`);
            }
            
        } catch (error) {
            console.error('Error parsing AI response:', error);
            // Fallback: display raw response and show relevant data
            this.addMessage(`AI: ${aiResponse}`, 'bot');
            const fallbackData = this.getSmartFilteredData(originalQuery);
            this.displayDataTable(fallbackData, `Results for: "${originalQuery}" (${fallbackData.length} rows)`);
        }
    }

    getSmartFilteredData(query) {
        if (!this.excelData || this.excelData.length === 0) return [];
        
        const lowerQuery = query.toLowerCase();
        
        // Show all data for broad queries
        if (lowerQuery.includes('all data') || lowerQuery.includes('show all') || 
            lowerQuery.includes('complete data') || lowerQuery.includes('full dataset')) {
            return this.excelData.slice(0, 50); // Limit to 50 rows
        }
        
        // Show first N rows
        if (lowerQuery.includes('first')) {
            const match = lowerQuery.match(/first\s+(\d+)/);
            const count = match ? Math.min(parseInt(match[1]), 20) : 10;
            return this.excelData.slice(0, count);
        }
        
        // Show sample
        if (lowerQuery.includes('sample') || lowerQuery.includes('example')) {
            return this.excelData.slice(0, 10);
        }
        
        // Filter by specific column values
        for (const header of this.headers) {
            if (lowerQuery.includes(header.toLowerCase())) {
                // Try to find rows that match the query in this column
                const filtered = this.excelData.filter(row => {
                    const value = row[header]?.toString().toLowerCase();
                    return value && lowerQuery.includes(value);
                });
                if (filtered.length > 0) return filtered.slice(0, 20);
            }
        }
        
        // Default: return first 15 rows
        return this.excelData.slice(0, 15);
    }

    displayDataTable(data, title = 'Data Results') {
        const tableWrapper = document.getElementById('tableWrapper');
        
        if (!data || data.length === 0) {
            tableWrapper.innerHTML = '<div class="status-message">No data to display</div>';
            this.updateTableInfo('No data available');
            return;
        }

        try {
            // Update table info
            this.updateTableInfo(title);
            
            // Create table
            let tableHTML = '<table><thead><tr>';
            
            // Create headers from the first data row
            const headers = Object.keys(data[0]);
            headers.forEach(header => {
                tableHTML += `<th>${this.escapeHtml(header)}</th>`;
            });
            tableHTML += '</tr></thead><tbody>';

            // Create rows
            data.forEach(row => {
                tableHTML += '<tr>';
                headers.forEach(header => {
                    const value = row[header];
                    tableHTML += `<td>${this.escapeHtml(value !== null && value !== undefined ? value.toString() : '')}</td>`;
                });
                tableHTML += '</tr>';
            });

            tableHTML += '</tbody></table>';
            tableWrapper.innerHTML = tableHTML;
            
        } catch (error) {
            console.error('Error displaying table:', error);
            tableWrapper.innerHTML = '<div class="status-message">Error displaying data table</div>';
            this.updateTableInfo('Error displaying data');
        }
    }

    escapeHtml(unsafe) {
        if (unsafe === null || unsafe === undefined) return '';
        return unsafe.toString()
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new ExcelAIQuery();
});