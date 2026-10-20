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

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
        
        btn.classList.add('active');
        document.getElementById(`${btn.dataset.tab}-tab`).classList.add('active');
    });
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
        if (pdf.size > 300 * 1024 * 1024) {
            showError('PDF file must be under 300MB');
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
    } else if (activeTab === 'text') {
        formData.append('text', document.getElementById('text').value.trim());
    }
    
    try {
        const response = await fetch('/api/generate', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.detail || 'Generation failed');
        }
        
        currentFilename = data.filename;
        showResult(data);
    } catch (err) {
        showError(err.message);
    } finally {
        setLoading(false);
    }
});

downloadBtn.addEventListener('click', () => {
    if (currentFilename) {
        window.location.href = `/api/download/${currentFilename}`;
    }
});

newBtn.addEventListener('click', () => {
    form.reset();
    document.querySelector('.tab-btn[data-tab="url"]').click();
    hideResult();
    hideError();
    currentFilename = '';
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