/// Ensure Office is ready before using the API
Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        const settingsButton = document.getElementById('settingsButton');
        const settingsModal = document.getElementById('settingsModal');
        const closeSettings = document.getElementById('closeSettings');
        const settingsForm = document.getElementById('settingsForm');
        const classificationForm = document.getElementById('classificationForm');
        const selectRangeButton = document.getElementById('selectRangeButton');
        const selectCategoriesButton = document.getElementById('selectCategoriesButton');
        const notification = document.getElementById('notification');
        const notificationMessage = document.getElementById('notificationMessage');
        const resultContainer = document.getElementById('result');
        const fixedModel = 'gpt-4o-mini';
        const fixedMaxTokens = 50;
        const fixedTemperature = 0.5;

        // Utility function to show error or success messages
        function showNotification(message, isError = true, autoHide = true) {
            notificationMessage.textContent = message;
            notification.style.backgroundColor = isError ? '#f8d7da' : '#d4edda';
            notification.style.color = isError ? '#721c24' : '#155724';
            notification.style.display = 'block';

            if (autoHide) {
                setTimeout(() => {
                    notification.style.display = 'none';
                }, 5000);
            }
        }

        function hideNotification() {
            notification.style.display = 'none';
        }

        // Mask API key for display
        const maskApiKey = (apiKey) => {
            if (!apiKey || apiKey.length <= 4) return '****';
            return '********' + apiKey.slice(-4);
        };

        // Load saved settings from localStorage
        const loadSettings = () => {
            const storedApiKey = localStorage.getItem('apiKey');
            document.getElementById('apiKey').value = storedApiKey ? maskApiKey(storedApiKey) : '';
        };

        // Save settings to localStorage
        const saveSettings = () => {
            const apiKeyInput = document.getElementById('apiKey').value.trim();
            const storedApiKey = localStorage.getItem('apiKey');

            // Check if the user has entered a new key or kept the masked one
            if (apiKeyInput && apiKeyInput !== maskApiKey(storedApiKey)) {
                localStorage.setItem('apiKey', apiKeyInput); // Save the new API key
            }
            showNotification('Settings saved successfully!', false);
            return true;
        };

        // Open settings modal
        settingsButton.onclick = () => {
            loadSettings();
            settingsModal.style.display = 'block';
        };

        // Close settings modal
        closeSettings.onclick = () => {
            settingsModal.style.display = 'none';
        };

        // Save settings and close modal
        settingsForm.onsubmit = (e) => {
            e.preventDefault();
            if (saveSettings()) {
                settingsModal.style.display = 'none';
            }
        };

        // Select input range from Excel
        if (selectRangeButton) {
            selectRangeButton.onclick = async () => {
                try {
                    await Excel.run(async (context) => {
                        const range = context.workbook.getSelectedRange();
                        range.load(['address', 'values']);
                        await context.sync();

                        document.getElementById('inputRange').value = range.address;
                        document.getElementById('inputRange').dataset.values = JSON.stringify(range.values);
                        console.log('Selected input range:', range.values);
                        showNotification('Range selected successfully!', false);
                    });
                } catch (error) {
                    console.error('Error selecting range:', error);
                    showNotification('Error selecting range: ' + error.message, true);
                }
            };
        }

        // Select categories from Excel
        if (selectCategoriesButton) {
            selectCategoriesButton.onclick = async () => {
                try {
                    await Excel.run(async (context) => {
                        const range = context.workbook.getSelectedRange();
                        range.load(['address', 'values']);
                        await context.sync();

                        document.getElementById('categories').value = range.address;
                        document.getElementById('categories').dataset.values = JSON.stringify(range.values);
                        console.log('Selected categories:', range.values);
                        showNotification('Categories selected successfully!', false);
                    });
                } catch (error) {
                    console.error('Error selecting categories:', error);
                    showNotification('Error selecting categories: ' + error.message, true);
                }
            };
        }

        // Handle classification form submission
        classificationForm.onsubmit = async (e) => {
            e.preventDefault();

            const apiKey = localStorage.getItem('apiKey');
            const model = fixedModel;
            const maxTokens = fixedMaxTokens;
            const temperature = fixedTemperature;

            if (!apiKey) {
                showNotification('Please set your API key in the settings.', true);
                return;
            }

            const inputRangeAddress = document.getElementById('inputRange').value;
            const categoriesAddress = document.getElementById('categories').value;
            const instructions = document.getElementById('instructions').value;

            const inputRangeValues = JSON.parse(document.getElementById('inputRange').dataset.values || '[]');
            const categoriesValues = JSON.parse(document.getElementById('categories').dataset.values || '[]');

            if (!inputRangeAddress || !categoriesAddress) {
                showNotification('Please fill in all required fields!', true);
                return;
            }

            try {
                const payload = {
                    apiKey,
                    model,
                    maxTokens,
                    temperature,
                    inputRange: inputRangeAddress,
                    inputData: inputRangeValues,
                    categories: categoriesValues.flat(),
                    instructions
                };

                console.log('Payload being sent to backend:', payload);

                // Show "Processing your request..." notification without auto-hide
                showNotification('Processing your request...', false, false);

                const response = await fetch('http://localhost:5000/api/analyze', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const result = await response.json();
                if (response.ok) {
                    console.log('Response from backend:', result);

                    // Prepare data for the new sheet
                    const resultData = result.results.map((item) => [
                        item.item,
                        item.category,
                        parseFloat(item.probability.toFixed(2)), // Format probability to two decimal places
                    ]);

                    console.log('Prepared data for Excel:', resultData);

                    // Write data to a new worksheet in Excel
                    await Excel.run(async (context) => {
                        const workbook = context.workbook;

                        // Check if the "Analysis Results" worksheet already exists
                        let analysisSheet = workbook.worksheets.getItemOrNullObject('Analysis Results');
                        await context.sync();

                        if (!analysisSheet.isNullObject) {
                            // Clear existing content if the sheet exists
                            analysisSheet.getUsedRange().clear();
                        } else {
                            // Create a new worksheet if it doesn't exist
                            analysisSheet = workbook.worksheets.add('Analysis Results');
                        }

                        // Prepare the header and data
                        const header = [['Item', 'Category', 'Probability']];
                        analysisSheet.getRange('A1:C1').values = header;
                        analysisSheet.getRange(`A2:C${resultData.length + 1}`).values = resultData;

                        // Activate the worksheet
                        analysisSheet.activate();

                        // Sync changes
                        await context.sync();
                        showNotification('Results updated in the "Analysis Results" sheet!', false);
                    });
                } else {
                    console.error('Error from server:', result.error);
                    showNotification('Error: ' + result.error, true);
                }
            } catch (err) {
                console.error('Error during fetch:', err);
                showNotification('An error occurred while connecting to the server.', true);
            } finally {
                // Hide the "Processing your request..." notification
                hideNotification();
            }
        };

        // Initialize the app and check if settings are saved
        const initializeApp = () => {
            if (!localStorage.getItem('apiKey')) {
                settingsModal.style.display = 'block';
            }
        };

        initializeApp();
    }
});
