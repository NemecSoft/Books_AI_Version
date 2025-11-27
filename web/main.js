// 从data.js引入数据 - 数据已移至单独文件以提高可维护性

// 使用全局变量的加载函数
let cardData = [];

async function loadData() {
    try {
        // 从全局变量获取数据，该变量在data.js中定义
        cardData = globalThis.xiyoujiData || [];
        return cardData; // 返回数据
    } catch (error) {
        console.error('加载数据失败:', error);
        // 如果出现意外错误，使用备用数据
        cardData = [
            {
                id: 1,
                name: "数据加载失败",
                description: "请检查数据文件是否存在或网络连接是否正常。",
                imageCount: 1,
                path: ""
            }
        ];
        return cardData;
    }
}

// 创建更美观的占位符SVG
function createPlaceholderSVG(cardId, imageIndex, storyName = '') {
    const storyNumber = cardId;
    const imageNumber = String(imageIndex).padStart(2, '0');
    
    // 如果有故事名称，则添加到SVG中
    let nameElement = '';
    if (storyName && storyName.length > 0) {
        // 对故事名称进行编码，确保SVG中可以正确显示
        const encodedName = encodeURIComponent(storyName);
        nameElement = `<text x="400" y="590" id="storyName" text-anchor="middle" fill="%234A5568" font-weight="normal" font-family="Helvetica%2C%20Arial%2C%20sans-serif" font-size="18pt">${encodedName}</text>`;
    }
    
    return `data:image/svg+xml;charset=UTF-8,%3Csvg%20width%3D%22150%22%20height%3D%22150%22%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20viewBox%3D%220%200%20800%20800%22%20preserveAspectRatio%3D%22none%22%3E%3Cdefs%3E%3Cstyle%20type%3D%22text%2Fcss%22%3E%23background%20%7B%20fill%3A%23F0F4F8%3B%7D%20%23border%20%7B%20fill%3Anone%3Bstroke%3A%23D1D9E6%3Bstroke-width%3A4%3B%7D%20%23icon%20%7B%20fill%3A%23A0AEC0%3B%7D%20%23storyNumber%20%7B%20fill%3A%234A5568%3Bfont-weight%3Abold%3Bfont-family%3AHelvetica%2C%20Arial%2C%20sans-serif%3Bfont-size%3A32pt%20%7D%20%23imageNumber%20%7B%20fill%3A%23718096%3Bfont-weight%3Anormal%3Bfont-family%3AHelvetica%2C%20Arial%2C%20sans-serif%3Bfont-size%3A24pt%20%7D%20%3C%2Fstyle%3E%3C%2Fdefs%3E%3Crect%20id%3D%22background%22%20width%3D%22800%22%20height%3D%22800%22%3E%3C%2Frect%3E%3Crect%20id%3D%22border%22%20x%3D%2210%22%20y%3D%2210%22%20width%3D%22780%22%20height%3D%22780%22%20rx%3D%2215%22%3E%3C%2Frect%3E%3Cg%20transform%3D%22translate%28350%2C%20300%29%22%3E%3Cpath%20id%3D%22icon%22%20d%3D%22M40%2C20%20H10%20C5%2C20%201%2C24%201%2C29%20V54%20C1%2C59%205%2C63%2010%2C63%20H14%20L20%2C69%20L26%2C63%20H60%20C65%2C63%2069%2C59%2069%2C54%20V29%20C69%2C24%2065%2C20%2060%2C20%20Z%22%20stroke%3D%22%23CBD5E0%22%20stroke-width%3D%222%22%3E%3C%2Fpath%3E%3Cpath%20d%3D%22M27%2C35%20H53%20C55%2C35%2057%2C37%2057%2C39%20V50%20C57%2C52%2055%2C54%2053%2C54%20H27%20C25%2C54%2023%2C52%2023%2C50%20V39%20C23%2C37%2025%2C35%2027%2C35%20Z%22%20fill%3D%22%23E2E8F0%22%3E%3C%2Fpath%3E%3Ccircle%20cx%3D%2217%22%20cy%3D%2237%22%20r%3D%223%22%20fill%3D%22%23E2E8F0%22%3E%3C%2Fcircle%3E%3C%2Fg%3E%3Ctext%20x%3D%22400%22%20y%3D%22450%22%20id%3D%22storyNumber%22%20text-anchor%3D%22middle%22%3E故事%20${storyNumber}%3C%2Ftext%3E%3Ctext%20x%3D%22400%22%20y%3D%22520%22%20id%3D%22imageNumber%22%20text-anchor%3D%22middle%22%3E图${imageNumber}%3C%2Ftext%3E${nameElement}%3C%2Fsvg%3E`;
}

// 缩略图初始化函数 - 优化以支持文件系统直接打开
function initThumbnails(cardId, images) {
    const grid = document.getElementById(`thumbnailGrid${cardId}`);
    if (!grid) {
        console.warn(`缩略图容器 thumbnailGrid${cardId} 不存在`);
        return;
    }
    
    // 清空现有内容，避免重复添加
    grid.innerHTML = '';
    
    // 获取对应的卡片数据
    const card = cardData.find(c => c.id === cardId);
    const storyName = card ? (card.title || card.name) : '';
    
    console.log(`初始化卡片 ${cardId} (${storyName}) 的缩略图，共 ${images.length} 张图片`);
    
    images.forEach((src, index) => {
        const item = document.createElement('div');
        item.className = 'thumbnail-item';
        
        const img = document.createElement('img');
        img.src = src;
        img.alt = `图片 ${index + 1}`;
        
        // 增强的错误处理 - 记录详细的错误信息
        img.onerror = function() {
            console.log(`图片加载失败: ${src}，显示占位符`);
            // 图片加载失败时显示占位符，并传入故事名称
            this.src = createPlaceholderSVG(cardId, index + 1, storyName);
            this.alt = `占位图 ${index + 1}`;
            // 为占位符添加特殊样式
            this.classList.add('placeholder-image');
        };
        
        // 添加加载成功事件处理
        img.onload = function() {
            console.log(`图片加载成功: ${src}`);
            // 加载成功时移除可能存在的占位符样式
            this.classList.remove('placeholder-image');
        };
        
        // 添加点击事件
        item.addEventListener('click', () => {
            // 这里可以添加查看大图的逻辑
            console.log(`查看卡片 ${cardId} 的图片 ${index + 1}`);
        });
        
        item.appendChild(img);
        grid.appendChild(item);
    });
}

// 动态创建卡片元素
function createCard(card) {
    const cardElement = document.createElement('div');
    cardElement.className = 'card';
    
    const cardHeader = document.createElement('div');
    cardHeader.className = 'card-header';
    
    const cardTitle = document.createElement('div');
        cardTitle.className = 'card-title';
        // 在标题前添加序号，格式为：序号. 标题
        cardTitle.textContent = `${card.id}. ${card.title || card.name}`;
    
    const cardDescription = document.createElement('div');
    cardDescription.className = 'card-description';
    cardDescription.textContent = card.description;
    
    cardHeader.appendChild(cardTitle);
    cardHeader.appendChild(cardDescription);
    
    const thumbnailGrid = document.createElement('div');
    thumbnailGrid.className = 'thumbnail-grid';
    thumbnailGrid.id = `thumbnailGrid${card.id}`;
    
    cardElement.appendChild(cardHeader);
    cardElement.appendChild(thumbnailGrid);
    
    return cardElement;
}

// 初始化所有卡片
function initCards() {
    const container = document.querySelector('.container');
    
    cardData.forEach(card => {
        // 创建卡片元素
        const cardElement = createCard(card);
        container.appendChild(cardElement);
        
        // 初始化卡片内容
        initCard(card);
    });
}

// 初始化单个卡片
function initCard(card) {
    // 生成图片路径数组
    const images = [];
    
    // 为每个应该显示的图片创建路径
    for (let i = 1; i <= card.imageCount; i++) {
        const paddedIndex = String(i).padStart(2, '0');
        
        // 构建图片路径 - 直接使用相对路径，确保在文件系统中直接打开时也能工作
        // 优先使用card.path，如果没有则使用card.id作为目录名
        const imagePath = card.path ? `${card.path}/${paddedIndex}.png` : `${card.id}/${paddedIndex}.png`;
        
        // 记录构建的路径，便于调试
        console.log(`构建图片路径: ${imagePath} 用于卡片 ${card.title || card.name}`);
        
        images.push(imagePath);
    }

    // 初始化缩略图
    initThumbnails(card.id, images);
}

// 初始化
// 移除多余的DOMContentLoaded事件监听器，避免重复加载
// 主初始化逻辑在HTML文件中的单个DOMContentLoaded事件处理程序中执行