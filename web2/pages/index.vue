<template>
  <div class="home-container">
    <h2>欢迎来到创新小说阅读平台</h2>
    
    <section class="feature-section">
      <h3>平台特色</h3>
      <ul>
        <li>支持小说原文阅读</li>
        <li>提供事件列表和简要事件</li>
        <li>AI绘图和AI视频支持</li>
        <li>用户分类权限管理</li>
        <li>付费内容和广告系统</li>
      </ul>
    </section>
    
    <section class="popular-section">
      <h3>热门小说</h3>
      
      <!-- 加载状态 -->
      <div v-if="loading" class="loading-container">
        <div class="loading-spinner"></div>
        <p>加载中...</p>
      </div>
      
      <!-- 错误信息 -->
      <div v-else-if="error" class="error-container">
        <p class="error-message">{{ error }}</p>
        <button @click="fetchNovels" class="btn-secondary">重试</button>
      </div>
      
      <!-- 小说列表 -->
      <div v-else class="novel-list">
        <div v-if="novels.length === 0" class="empty-state">
          <p>暂无小说数据</p>
        </div>
        <div v-else v-for="novel in novels" :key="novel.id" class="novel-card">
          <NuxtLink :to="`/novels/${novel.id}`" class="novel-link">
            <h4>{{ novel.title }}</h4>
            <p>{{ novel.author }}</p>
            <span v-if="novel.is_paid" class="paid-badge">付费</span>
          </NuxtLink>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue';

const novels = ref<any[]>([]);
const loading = ref(true);
const error = ref('');

// 硬编码的模拟数据
const mockNovels = [
  {
    id: 1,
    title: '星际漫游',
    author: '刘慈欣',
    category: '科幻',
    description: '一部关于星际旅行的科幻小说，讲述了人类探索宇宙的壮丽征程。',
    cover: 'https://example.com/cover1.jpg',
    is_paid: true,
    createdAt: new Date().toISOString()
  },
  {
    id: 2,
    title: '魔法世界',
    author: 'JK罗琳',
    category: '奇幻',
    description: '一个充满魔法的奇幻世界，勇敢的小魔法师踏上冒险之旅。',
    cover: 'https://example.com/cover2.jpg',
    is_paid: false,
    createdAt: new Date().toISOString()
  },
  {
    id: 3,
    title: '战争与和平',
    author: '托尔斯泰',
    category: '文学',
    description: '一部经典的文学作品，描绘了战争年代的人性与爱情。',
    cover: 'https://example.com/cover3.jpg',
    is_paid: true,
    createdAt: new Date().toISOString()
  },
  {
    id: 4,
    title: '百年孤独',
    author: '加西亚·马尔克斯',
    category: '魔幻现实主义',
    description: '魔幻现实主义文学的代表作，讲述了布恩迪亚家族七代人的传奇故事。',
    cover: 'https://example.com/cover4.jpg',
    is_paid: false,
    createdAt: new Date().toISOString()
  }
];

// 使用硬编码数据替代API调用
const fetchNovels = async () => {
  loading.value = true;
  error.value = '';
  
  try {
    // 直接使用模拟数据，不再调用API
    console.log('使用硬编码模拟数据');
    novels.value = mockNovels;
  } catch (err) {
    error.value = '加载失败';
    console.error('获取小说列表失败:', err);
  } finally {
    loading.value = false;
  }
};

onMounted(() => {
  fetchNovels();
});
</script>

<style scoped>
.home-container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 20px;
}

h2 {
  margin-bottom: 30px;
  color: #333;
}

.feature-section {
  margin-bottom: 40px;
  padding: 20px;
  background-color: #f9f9f9;
  border-radius: 8px;
}

.feature-section ul {
  list-style: none;
  padding: 0;
}

.feature-section li {
  padding: 8px 0;
  padding-left: 20px;
  position: relative;
}

.feature-section li:before {
  content: '✦';
  position: absolute;
  left: 0;
  color: #4CAF50;
}

.popular-section h3 {
  margin-bottom: 20px;
}

.novel-list {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 20px;
}

.novel-card {
  background-color: #f5f5f5;
  padding: 20px;
  border-radius: 8px;
  transition: transform 0.3s, box-shadow 0.3s;
  position: relative;
}

.novel-card:hover {
  transform: translateY(-5px);
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

.novel-link {
  text-decoration: none;
  color: #333;
}

.novel-link h4 {
  margin-bottom: 10px;
  color: #2196F3;
}

/* 加载状态样式 */
.loading-container {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 50px 0;
}

.loading-spinner {
  width: 50px;
  height: 50px;
  border: 4px solid #f3f3f3;
  border-top: 4px solid #667eea;
  border-radius: 50%;
  animation: spin 1s linear infinite;
  margin-bottom: 15px;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

/* 错误信息样式 */
.error-container {
  text-align: center;
  padding: 50px 0;
}

.error-message {
  color: #e53e3e;
  margin-bottom: 20px;
  font-size: 1.1rem;
}

.btn-secondary {
  background-color: #e2e8f0;
  color: #4a5568;
  border: none;
  padding: 10px 20px;
  border-radius: 5px;
  cursor: pointer;
  font-size: 1rem;
  transition: background-color 0.2s;
}

.btn-secondary:hover {
  background-color: #cbd5e0;
}

/* 空状态样式 */
.empty-state {
  text-align: center;
  padding: 50px 0;
  color: #718096;
  font-size: 1.1rem;
}

/* 付费标签样式 */
.paid-badge {
  position: absolute;
  top: 10px;
  right: 10px;
  background-color: #e53e3e;
  color: white;
  padding: 3px 8px;
  border-radius: 4px;
  font-size: 0.8rem;
  font-weight: bold;
  z-index: 1;
}
</style>