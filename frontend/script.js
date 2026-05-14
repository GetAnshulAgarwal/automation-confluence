const compareBtn = document.getElementById('compareBtn');
const statusDiv  = document.getElementById('status');

function setStatus(message, type) {
    statusDiv.innerText = message;
    statusDiv.className = type;
}

compareBtn.addEventListener('click', async () => {

    const docFile     = document.getElementById('docFile').files[0];
    const folderFiles = document.getElementById('folderInput').files;

    if (!docFile) {
        alert('Please upload a PDF or DOCX file');
        return;
    }

    if (!folderFiles || folderFiles.length === 0) {
        alert('Please upload the app folder');
        return;
    }

    const formData = new FormData();
    formData.append('document', docFile);

    for (let file of folderFiles) {
        formData.append('yamlFiles', file, file.webkitRelativePath);
    }

    try {
        compareBtn.disabled = true;
        setStatus('⏳ Comparing... please wait', 'info');

        const response = await fetch('http://127.0.0.1:5000/compare', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            let errMsg = 'Comparison failed';
            try {
                const errJson = await response.json();
                errMsg = errJson.error || errMsg;
            } catch (_) {}
            throw new Error(errMsg);
        }

        const blob = await response.blob();
        const url  = window.URL.createObjectURL(blob);
        const a    = document.createElement('a');

        a.href     = url;
        a.download = 'comparison_result.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();

        setStatus('✅ Excel downloaded successfully! Open the file to see mismatches.', 'success');

    } catch (error) {
        console.error(error);
        setStatus('❌ Error: ' + error.message, 'error');

    } finally {
        compareBtn.disabled = false;
    }
});