<template>
  <div class="novel-detail-container">
    <div v-if="loading" class="loading">加载中...</div>
    
    <div v-else-if="novel" class="novel-detail">
      <div class="novel-header">
        <div class="novel-cover-large">
          <img :src="novel.cover_image || '/default-cover.png'" :alt="novel.title" />
        </div>
        <div class="novel-main-info">
          <h1>{{ novel.title }}</h1>
          <p class="author">作者：{{ novel.author }}</p>
          <div class="meta-info">
            <span class="category">分类：{{ novel.category || '未分类' }}</span>
            <span v-if="novel.is_paid" class="paid-badge">付费小说</span>
          </div>
          <div class="action-buttons">
            <button @click="startReading" class="primary-button">开始阅读</button>
            <button @click="addFavorite" class="secondary-button">收藏</button>
          </div>
        </div>
      </div>
      
      <div class="novel-description">
        <h2>作品简介</h2>
        <p>{{ novel.description || '暂无简介' }}</p>
      </div>
      
      <div class="novel-chapters">
        <h2>章节列表</h2>
        <div class="chapters-list">
          <div v-for="chapter in chapters" :key="chapter.id" class="chapter-item">
            <NuxtLink 
              :to="`/novels/${novel.id}/chapters/${chapter.id}`" 
              class="chapter-link"
            >
              <span class="chapter-number">{{ chapter.chapter_number }}. </span>
              <span class="chapter-title">{{ chapter.title }}</span>
              <span v-if="chapter.is_paid" class="chapter-paid">[付费]</span>
            </NuxtLink>
          </div>
        </div>
      </div>
      
      <div class="events-section">
        <h2>事件列表</h2>
        <div class="events-list">
          <div v-for="event in events" :key="event.id" class="event-item">
            <div class="event-time">{{ event.timestamp }}</div>
            <div class="event-text">{{ event.text }}</div>
          </div>
        </div>
      </div>
    </div>
    
    <div v-else class="error">小说不存在</div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()
const novelId = Number(route.params.id)

// 状态变量
const novel = ref(null)
const chapters = ref([])
const events = ref([])
const loading = ref(true)

// 模拟数据
const mockNovels = [
  {
    id: 1,
    title: '星际迷航',
    author: '张小明',
    description: '一部关于星际探险的科幻小说，主角穿越星系寻找新家园的故事。在遥远的未来，人类面临着地球资源枯竭的危机，一群勇敢的探险家踏上了寻找新家园的征程。在旅途中，他们遇到了各种奇异的外星生物和文明，也经历了无数的危险和挑战。',
    cover_image: '',
    category: '科幻',
    is_paid: false
  },
  {
    id: 2,
    title: '魔法学院',
    author: '李华',
    description: '讲述了一个普通少年进入魔法学院学习的奇幻冒险故事。少年亚瑟在12岁生日那天收到了来自霍格沃茨魔法学院的录取通知书，从此踏上了学习魔法的奇幻旅程。',
    cover_image: '',
    category: '奇幻',
    is_paid: true
  }
]

const mockChapters = {
  1: [
    { id: 1, novel_id: 1, title: '序章：地球末日', chapter_number: 0, is_paid: false },
    { id: 2, novel_id: 1, title: '起航', chapter_number: 1, is_paid: false },
    { id: 3, novel_id: 1, title: '深空探索', chapter_number: 2, is_paid: false },
    { id: 4, novel_id: 1, title: '第一次接触', chapter_number: 3, is_paid: true }
  ],
  2: [
    { id: 5, novel_id: 2, title: '奇怪的信件', chapter_number: 0, is_paid: false },
    { id: 6, novel_id: 2, title: '魔法学院', chapter_number: 1, is_paid: false },
    { id: 7, novel_id: 2, title: '分院仪式', chapter_number: 2, is_paid: true }
  ]
}

const mockEvents = [
  { id: 1, text: '地球资源枯竭，人类开始寻找新家园', timestamp: '2150年' },
  { id: 2, text: '星际飞船「希望号」发射成功', timestamp: '2152年3月' },
  { id: 3, text: '发现第一个可居住行星', timestamp: '2155年7月' },
  { id: 4, text: '与外星文明首次接触', timestamp: '2157年1月' }
]

onMounted(async () => {
  try {
    // 模拟API请求延迟
    await new Promise(resolve => setTimeout(resolve, 500))
    
    // 获取小说信息
    novel.value = mockNovels.find(n => n.id === novelId)
    
    // 获取章节列表
    if (novel.value) {
      chapters.value = mockChapters[novelId] || []
      events.value = mockEvents
    }
  } catch (error) {
    console.error('加载小说信息失败:', error)
  } finally {
    loading.value = false
  }
})

// 开始阅读
const startReading = () => {
  if (chapters.value.length > 0) {
    const firstChapter = chapters.value[0]
    location.href = `/novels/${novelId}/chapters/${firstChapter.id}`
  }
}

// 收藏小说
const addFavorite = () => {
  alert('收藏成功！')
}
</script>

<style scoped>
.novel-detail-container {
  padding: 20px;
}

.loading, .error {
  text-align: center;
  padding: 40px;
  font-size: 16px;
}

.error {
  color: #f44336;
}

.novel-header {
  display: flex;
  margin-bottom: 30px;
  gap: 30px;
  background-color: white;
  padding: 20px;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.novel-cover-large {
  width: 200px;
  height: 300px;
  overflow: hidden;
  background-color: #f0f0f0;
  border-radius: 8px;
  flex-shrink: 0;
}

.novel-cover-large img {
  width: 100%;
  height: 100%;
  object-fit: cover;
}

.novel-main-info {
  flex: 1;
}

.novel-main-info h1 {
  margin-bottom: 10px;
  font-size: 28px;
}

.author {
  font-size: 18px;
  color: #666;
  margin-bottom: 15px;
}

.meta-info {
  display: flex;
  gap: 20px;
  margin-bottom: 20px;
  font-size: 14px;
}

.category {
  color: #666;
}

.paid-badge {
  background-color: #ff6b6b;
  color: white;
  padding: 4px 8px;
  border-radius: 10px;
}

.action-buttons {
  display: flex;
  gap: 15px;
}

.primary-button {
  background-color: #2196F3;
  color: white;
  padding: 10px 20px;
  border: none;
  border-radius: 4px;
  font-size: 16px;
  cursor: pointer;
}

.secondary-button {
  background-color: transparent;
  color: #2196F3;
  padding: 10px 20px;
  border: 1px solid #2196F3;
  border-radius: 4px;
  font-size: 16px;
  cursor: pointer;
}

.novel-description {
  background-color: white;
  padding: 20px;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  margin-bottom: 30px;
}

.novel-description h2 {
  margin-bottom: 15px;
  font-size: 20px;
}

.novel-description p {
  line-height: 1.8;
  color: #333;
}

.novel-chapters {
  background-color: white;
  padding: 20px;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  margin-bottom: 30px;
}

.novel-chapters h2 {
  margin-bottom: 20px;
  font-size: 20px;
}

.chapters-list {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  gap: 10px;
}

.chapter-item {
  padding: 10px;
  border-bottom: 1px solid #eee;
}

.chapter-link {
  display: flex;
  align-items: center;
  text-decoration: none;
  color: #333;
  transition: color 0.3s;
}

.chapter-link:hover {
  color: #2196F3;
}

.chapter-number {
  font-weight: bold;
  margin-right: 5px;
}

.chapter-title {
  flex: 1;
}

.chapter-paid {
  color: #ff6b6b;
  font-size: 12px;
}

.events-section {
  background-color: white;
  padding: 20px;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.events-section h2 {
  margin-bottom: 20px;
  font-size: 20px;
}

.events-list {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.event-item {
  display: flex;
  gap: 15px;
  padding: 10px;
  border-left: 3px solid #2196F3;
  background-color: #f9f9f9;
}

.event-time {
  font-weight: bold;
  color: #2196F3;
  min-width: 100px;
}

.event-text {
  flex: 1;
}

@media (max-width: 768px) {
  .novel-header {
    flex-direction: column;
    align-items: center;
    text-align: center;
  }
  
  .novel-cover-large {
    width: 150px;
    height: 220px;
  }
  
  .chapters-list {
    grid-template-columns: 1fr;
  }
}
</style>