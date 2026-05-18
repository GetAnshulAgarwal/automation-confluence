const API = 'http://127.0.0.1:5000';

const loadFoldersBtn        = document.getElementById('loadFoldersBtn');
const compareBtn            = document.getElementById('compareBtn');
const statusDiv             = document.getElementById('status');
const folderSelectionSection= document.getElementById('folderSelectionSection');
const folderList            = document.getElementById('folderList');
const selectedCount         = document.getElementById('selectedCount');
const selectAllBtn          = document.getElementById('selectAllBtn');
const deselectAllBtn        = document.getElementById('deselectAllBtn');

// ─── Helpers ────────────────────────────────────────────────
function setStatus(message, type) {
    statusDiv.innerText = message;
    statusDiv.className = type;
}

function updateCount() {
    const total    = folderList.querySelectorAll('input[type=checkbox]').length;
    const checked  = folderList.querySelectorAll('input[type=checkbox]:checked').length;
    selectedCount.textContent = `${checked} / ${total} selected`;
    compareBtn.style.display = checked > 0 ? 'block' : 'none';
}

// ─── Load Folders ────────────────────────────────────────────
loadFoldersBtn.addEventListener('click', async () => {
    const appFolderPath = document.getElementById('appFolderPath').value.trim().replace(/^["']+|["']+$/g, '');

    if (!appFolderPath) {
        alert('Please enter the App Folder Path first');
        return;
    }

    setStatus('', '');
    loadFoldersBtn.disabled = true;
    loadFoldersBtn.textContent = 'Loading...';

    try {
        const res = await fetch(`${API}/get-folders`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ appFolderPath })
        });

        const data = await res.json();

        if (!res.ok) throw new Error(data.error || 'Failed to load folders');

        const folders = data.folders;

        if (!folders || folders.length === 0) {
            setStatus('❌ No subfolders found in the specified path.', 'error');
            return;
        }

        // Build checkbox grid
        folderList.innerHTML = '';
        folders.forEach(folder => {
            const label = document.createElement('label');
            label.className = 'folder-item';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = folder;
            checkbox.checked = true; // default: all selected
            checkbox.addEventListener('change', updateCount);

            const span = document.createElement('span');
            span.textContent = folder;

            label.appendChild(checkbox);
            label.appendChild(span);
            folderList.appendChild(label);
        });

        folderSelectionSection.style.display = 'block';
        updateCount();
        setStatus(`✅ ${folders.length} folders loaded. Select which to compare.`, 'success');

    } catch (err) {
        setStatus('❌ Error: ' + err.message, 'error');
    } finally {
        loadFoldersBtn.disabled = false;
        loadFoldersBtn.textContent = 'Load Folders';
    }
});

// ─── Select / Deselect All ───────────────────────────────────
selectAllBtn.addEventListener('click', () => {
    folderList.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = true);
    updateCount();
});

deselectAllBtn.addEventListener('click', () => {
    folderList.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
    updateCount();
});

// ─── Compare ─────────────────────────────────────────────────
compareBtn.addEventListener('click', async () => {

    const documentPath  = document.getElementById('documentPath').value.trim().replace(/^["']+|["']+$/g, '');
    const appFolderPath = document.getElementById('appFolderPath').value.trim().replace(/^["']+|["']+$/g, '');

    const selectedFolders = Array.from(
        folderList.querySelectorAll('input[type=checkbox]:checked')
    ).map(cb => cb.value);

    if (!documentPath) {
        alert('Please enter the Document Path');
        return;
    }

    if (selectedFolders.length === 0) {
        alert('Please select at least one folder');
        return;
    }

    try {
        compareBtn.disabled = true;
        setStatus(`⏳ Comparing ${selectedFolders.length} folder(s)... please wait`, 'info');

        const response = await fetch(`${API}/compare`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                documentPath,
                appFolderPath,
                selectedFolders
            })
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

        setStatus(`✅ Done! Excel downloaded with ${selectedFolders.length} folder(s) compared.`, 'success');

    } catch (err) {
        console.error(err);
        setStatus('❌ Error: ' + err.message, 'error');
    } finally {
        compareBtn.disabled = false;
    }
});