<template>
  <div class="novels-container">
    <h2>小说列表</h2>
    
    <div class="search-bar">
      <input
        v-model="searchKeyword"
        type="text"
        placeholder="搜索小说标题或作者..."
        @input="handleSearch"
      />
    </div>
    
    <div class="filter-section">
      <label for="category">分类筛选：</label>
      <select id="category" v-model="selectedCategory" @change="handleFilter">
        <option value="">全部分类</option>
        <option v-for="category in categories" :key="category" :value="category">
          {{ category }}
        </option>
      </select>
    </div>
    
    <div class="novel-grid">
      <div v-for="novel in displayedNovels" :key="novel.id" class="novel-card">
        <NuxtLink :to="`/novels/${novel.id}`" class="novel-link">
          <div class="novel-cover">
            <img :src="novel.cover_image || '/default-cover.png'" :alt="novel.title" />
          </div>
          <div class="novel-info">
            <h3>{{ novel.title }}</h3>
            <p class="author">{{ novel.author }}</p>
            <p class="description">{{ truncateDescription(novel.description) }}</p>
            <div class="novel-meta">
              <span v-if="novel.is_paid" class="paid-badge">付费小说</span>
              <span class="category">{{ novel.category || '未分类' }}</span>
            </div>
          </div>
        </NuxtLink>
      </div>
    </div>
    
    <div v-if="displayedNovels.length === 0" class="no-result">
      没有找到符合条件的小说
    </div>
    
    <div class="pagination" v-if="totalPages > 1">
      <button @click="currentPage > 1 && currentPage--" :disabled="currentPage <= 1">
        上一页
      </button>
      <span class="page-info">{{ currentPage }} / {{ totalPages }}</span>
      <button @click="currentPage < totalPages && currentPage++" :disabled="currentPage >= totalPages">
        下一页
      </button>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'

// 模拟数据
const mockNovels = [
  {
    id: 1,
    title: '星际迷航',
    author: '张小明',
    description: '一部关于星际探险的科幻小说，主角穿越星系寻找新家园的故事。',
    cover_image: '',
    category: '科幻',
    is_paid: false
  },
  {
    id: 2,
    title: '魔法学院',
    author: '李华',
    description: '讲述了一个普通少年进入魔法学院学习的奇幻冒险故事。',
    cover_image: '',
    category: '奇幻',
    is_paid: true
  },
  {
    id: 3,
    title: '未来战士',
    author: '王强',
    description: '在未来世界，人类与机器人的战争一触即发...',
    cover_image: '',
    category: '科幻',
    is_paid: false
  },
  {
    id: 4,
    title: '古代探案录',
    author: '赵敏',
    description: '古代侦探凭借智慧破解一个个离奇案件的故事集。',
    cover_image: '',
    category: '悬疑',
    is_paid: true
  },
  {
    id: 5,
    title: '都市爱情',
    author: '陈静',
    description: '现代都市中年轻人的爱情故事，充满欢笑与泪水。',
    cover_image: '',
    category: '言情',
    is_paid: false
  },
  {
    id: 6,
    title: '武侠传奇',
    author: '刘强',
    description: '江湖恩怨，侠义情仇，一部经典的武侠小说。',
    cover_image: '',
    category: '武侠',
    is_paid: true
  }
]

const novels = ref(mockNovels)
const searchKeyword = ref('')
const selectedCategory = ref('')
const currentPage = ref(1)
const pageSize = 10

// 分类列表
const categories = ref(['科幻', '奇幻', '悬疑', '言情', '武侠', '历史', '军事'])

// 计算显示的小说
const displayedNovels = computed(() => {
  let filtered = novels.value
  
  // 按关键词筛选
  if (searchKeyword.value) {
    const keyword = searchKeyword.value.toLowerCase()
    filtered = filtered.filter(novel => 
      novel.title.toLowerCase().includes(keyword) || 
      novel.author.toLowerCase().includes(keyword)
    )
  }
  
  // 按分类筛选
  if (selectedCategory.value) {
    filtered = filtered.filter(novel => novel.category === selectedCategory.value)
  }
  
  // 分页
  const startIndex = (currentPage.value - 1) * pageSize
  return filtered.slice(startIndex, startIndex + pageSize)
})

// 计算总页数
const totalPages = computed(() => {
  let filtered = novels.value
  
  if (searchKeyword.value) {
    const keyword = searchKeyword.value.toLowerCase()
    filtered = filtered.filter(novel => 
      novel.title.toLowerCase().includes(keyword) || 
      novel.author.toLowerCase().includes(keyword)
    )
  }
  
  if (selectedCategory.value) {
    filtered = filtered.filter(novel => novel.category === selectedCategory.value)
  }
  
  return Math.ceil(filtered.length / pageSize)
})

// 搜索处理
const handleSearch = () => {
  currentPage.value = 1
}

// 筛选处理
const handleFilter = () => {
  currentPage.value = 1
}

// 截断描述
const truncateDescription = (description) => {
  if (!description) return '暂无描述'
  return description.length > 100 ? description.substring(0, 100) + '...' : description
}
</script>

<style scoped>
.novels-container {
  padding: 20px;
}

.search-bar {
  margin-bottom: 20px;
}

.search-bar input {
  width: 100%;
  max-width: 500px;
  padding: 12px;
  font-size: 16px;
}

.filter-section {
  margin-bottom: 20px;
}

.filter-section select {
  padding: 8px;
  margin-left: 10px;
}

.novel-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  gap: 20px;
  margin-bottom: 30px;
}

.novel-card {
  background-color: white;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  transition: transform 0.3s, box-shadow 0.3s;
}

.novel-card:hover {
  transform: translateY(-5px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
}

.novel-link {
  text-decoration: none;
  color: inherit;
  display: block;
}

.novel-cover {
  height: 200px;
  overflow: hidden;
  background-color: #f0f0f0;
}

.novel-cover img {
  width: 100%;
  height: 100%;
  object-fit: cover;
}

.novel-info {
  padding: 15px;
}

.novel-info h3 {
  margin-bottom: 8px;
  font-size: 18px;
}

.author {
  color: #666;
  font-size: 14px;
  margin-bottom: 10px;
}

.description {
  color: #333;
  font-size: 14px;
  line-height: 1.5;
  margin-bottom: 10px;
}

.novel-meta {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 12px;
}

.paid-badge {
  background-color: #ff6b6b;
  color: white;
  padding: 2px 6px;
  border-radius: 10px;
}

.category {
  color: #666;
}

.pagination {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 20px;
  margin-top: 30px;
}

.page-info {
  font-size: 16px;
}

.no-result {
  text-align: center;
  padding: 40px;
  color: #666;
}

@media (max-width: 768px) {
  .novel-grid {
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
  }
}
</style>