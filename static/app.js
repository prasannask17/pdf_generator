const form = document.getElementById('generate-form');
const generateBtn = document.getElementById('generate-btn');
const btnText = generateBtn.querySelector('.btn-text');
const btnLoader = generateBtn.querySelector('.btn-loader');
const errorMessage = document.getElementById('error-message');
const result = document.getElementById('result');
const deckTitle = document.getElementById('deck-title');
const deckSubtitle = document.getElementById('deck-subtitle');
const slideCount = document.getElementById('slide-count');
const imageCount = document.getElementById('image-count');
const slidesPreview = document.getElementById('slides-preview');
const downloadBtn = document.getElementById('download-btn');
const newBtn = document.getElementById('new-btn');

let currentFilename = '';
let currentFormat = 'pptx'; // 'pptx' or 'pdf'

// Format toggle — PPT / PDF selector in the result section
document.querySelectorAll('.format-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentFormat = btn.dataset.format;
        updateDownloadButton();
    });
});

function updateDownloadButton() {
    if (currentFormat === 'pdf') {
        downloadBtn.textContent = 'Download .pdf';
    } else {
        downloadBtn.textContent = 'Download .pptx';
    }
}

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
        
        btn.classList.add('active');
        document.getElementById(`${btn.dataset.tab}-tab`).classList.add('active');

        // Disable inputs in inactive tabs so browser validation skips them.
        // Without this, a hidden <input type="url"> triggers "the string did
        // not match the expected pattern" when submitting the PDF tab.
        document.querySelectorAll('.tab-pane').forEach(p => {
            const active = p.id === `${btn.dataset.tab}-tab`;
            p.querySelectorAll('input, textarea').forEach(el => {
                if (active) {
                    el.removeAttribute('disabled');
                } else {
                    el.setAttribute('disabled', 'disabled');
                }
            });
        });
    });
});

// On page load, disable inputs in inactive tabs (URL is active by default)
document.querySelectorAll('.tab-pane:not(.active) input, .tab-pane:not(.active) textarea').forEach(el => {
    el.setAttribute('disabled', 'disabled');
});

form.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const topic = document.getElementById('topic').value.trim();
    const numSlides = document.getElementById('num_slides').value;
    const activeTab = document.querySelector('.tab-btn.active').dataset.tab;
    
    if (!topic) {
        showError('Please enter a presentation topic');
        return;
    }
    
    if (activeTab === 'url') {
        const url = document.getElementById('url').value.trim();
        if (!url) {
            showError('Please enter a URL');
            return;
        }
    } else if (activeTab === 'pdf') {
        const pdf = document.getElementById('pdf').files[0];
        if (!pdf) {
            showError('Please select a PDF file');
            return;
        }
        if (pdf.size > 100 * 1024 * 1024) {
            showError('PDF file must be under 100MB');
            return;
        }
    } else if (activeTab === 'text') {
        const text = document.getElementById('text').value.trim();
        if (!text || text.length < 100) {
            showError('Please paste at least 100 characters of content');
            return;
        }
    }
    
    setLoading(true);
    hideError();
    hideResult();
    
    const formData = new FormData();
    formData.append('topic', topic);
    formData.append('num_slides', numSlides);
    
    if (activeTab === 'url') {
        formData.append('url', document.getElementById('url').value.trim());
    } else if (activeTab === 'pdf') {
        formData.append('pdf', document.getElementById('pdf').files[0]);
        const pageStart = document.getElementById('page_start').value;
        const pageEnd = document.getElementById('page_end').value;
        if (pageStart) formData.append('page_start', pageStart);
        if (pageEnd) formData.append('page_end', pageEnd);
    } else if (activeTab === 'text') {
        formData.append('text', document.getElementById('text').value.trim());
    }
    
    try {
        const submitResponse = await fetch('/api/generate', {
            method: 'POST',
            body: formData,
        });

        const submitText = await submitResponse.text();
        if (!submitText) {
            throw new Error('Server returned an empty response. Please try again.');
        }

        let submitData;
        try {
            submitData = JSON.parse(submitText);
        } catch {
            throw new Error('Server returned an unexpected response. Please try again.');
        }

        if (!submitResponse.ok) {
            throw new Error(submitData.detail || 'Generation failed');
        }

        // Now poll for job status
        const jobId = submitData.job_id;
        if (!jobId) {
            throw new Error('No job ID returned from server.');
        }

        showProgress('Starting…');
        const result = await pollJobStatus(jobId);
        hideProgress();

        if (result.status === 'error') {
            throw new Error(result.error || 'Generation failed');
        }

        currentFilename = result.result.filename;
        showResult(result.result);
    } catch (err) {
        hideProgress();
        showError(err.message);
    } finally {
        setLoading(false);
    }
});

downloadBtn.addEventListener('click', () => {
    if (currentFilename) {
        if (currentFormat === 'pdf') {
            const pdfFilename = currentFilename.replace(/\.pptx$/, '.pdf');
            window.location.href = `/api/download-pdf/${currentFilename}`;
        } else {
            window.location.href = `/api/download/${currentFilename}`;
        }
    }
});

newBtn.addEventListener('click', () => {
    form.reset();
    document.querySelector('.tab-btn[data-tab="url"]').click();
    hideResult();
    hideError();
    currentFilename = '';
    currentFormat = 'pptx';
    document.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
    document.querySelector('.format-btn[data-format="pptx"]').classList.add('active');
    updateDownloadButton();
});

function showResult(data) {
    deckTitle.textContent = data.deck_title;
    deckSubtitle.textContent = data.subtitle || '';
    slideCount.textContent = `${data.slides.length} slides`;
    imageCount.textContent = `${data.image_count} images`;
    
    slidesPreview.innerHTML = '';
    data.slides.forEach((slide, i) => {
        const div = document.createElement('div');
        div.className = 'slide-item';
        div.innerHTML = `
            <h4>${i + 1}. ${escapeHtml(slide.title)}</h4>
            <ul>
                ${slide.bullets.map(b => `<li>${escapeHtml(b)}</li>`).join('')}
            </ul>
            ${slide.has_image ? '<span class="has-image">📷 Has image</span>' : ''}
        `;
        slidesPreview.appendChild(div);
    });
    
    result.classList.remove('hidden');
    result.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function hideResult() {
    result.classList.add('hidden');
}

function showError(message) {
    errorMessage.textContent = message;
    errorMessage.classList.remove('hidden');
}

function hideError() {
    errorMessage.classList.add('hidden');
}

function setLoading(loading) {
    generateBtn.disabled = loading;
    btnText.classList.toggle('hidden', loading);
    btnLoader.classList.toggle('hidden', !loading);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Job polling — avoids Render's 60s request timeout
async function pollJobStatus(jobId) {
    const maxAttempts = 120; // 120 x 3s = 6 min max
    for (let i = 0; i < maxAttempts; i++) {
        await new Promise(r => setTimeout(r, 3000)); // poll every 3 seconds
        try {
            const resp = await fetch(`/api/jobs/${jobId}`);
            const text = await resp.text();
            if (!text) continue;
            const job = JSON.parse(text);
            if (job.message) showProgress(job.message);
            if (job.status === 'done' || job.status === 'error') {
                return job;
            }
        } catch {
            // network hiccup — keep polling
        }
    }
    return { status: 'error', error: 'Timed out waiting for results. Please try again with fewer slides.' };
}

function showProgress(message) {
    let el = document.getElementById('progress-message');
    if (!el) {
        el = document.createElement('div');
        el.id = 'progress-message';
        el.className = 'progress-message';
        generateBtn.parentNode.insertBefore(el, generateBtn.nextSibling);
    }
    el.textContent = message + ' Please wait…';
    el.classList.remove('hidden');
}

function hideProgress() {
    const el = document.getElementById('progress-message');
    if (el) el.classList.add('hidden');
}