<template>
  <div class="chapter-reading-container">
    <div v-if="loading" class="loading">加载中...</div>
    
    <div v-else-if="chapter" class="chapter-reading">
      <div class="chapter-header">
        <h1>{{ chapter.title }}</h1>
        <p class="chapter-info">
          {{ novelTitle }} - 第{{ chapter.chapter_number }}章
          <span v-if="chapter.is_paid" class="paid-badge">[付费章节]</span>
        </p>
      </div>
      
      <div class="reading-settings">
        <div class="font-size-control">
          <button @click="decreaseFontSize" class="setting-btn">A-</button>
          <span class="current-size">{{ fontSize }}px</span>
          <button @click="increaseFontSize" class="setting-btn">A+</button>
        </div>
        
        <div class="theme-switch">
          <select v-model="currentTheme" @change="changeTheme">
            <option value="light">浅色模式</option>
            <option value="dark">深色模式</option>
            <option value="sepia">护眼模式</option>
          </select>
        </div>
      </div>
      
      <div class="chapter-content" :style="contentStyle">
        <div v-if="chapter.is_paid && !hasAccess" class="paid-content-blocked">
          <h3>此章节为付费内容</h3>
          <p>请购买后再阅读</p>
          <button @click="buyChapter" class="buy-button">立即购买</button>
        </div>
        <div v-else>
          <!-- 章节内容，这里使用模拟的段落 -->
          <p v-for="(paragraph, index) in chapterParagraphs" :key="index" class="content-paragraph">
            {{ paragraph }}
          </p>
          
          <!-- AI绘图内容 -->
          <div v-if="aiImages.length > 0" class="ai-content-section">
            <h3>AI绘图</h3>
            <div class="ai-images">
              <div v-for="image in aiImages" :key="image.id" class="ai-image-item">
                <img :src="image.url" :alt="image.description" />
                <p class="image-description">{{ image.description }}</p>
              </div>
            </div>
          </div>
          
          <!-- AI视频内容 -->
          <div v-if="aiVideos.length > 0" class="ai-content-section">
            <h3>AI视频</h3>
            <div class="ai-videos">
              <div v-for="video in aiVideos" :key="video.id" class="ai-video-item">
                <video :src="video.url" controls></video>
                <p class="video-description">{{ video.description }}</p>
              </div>
            </div>
          </div>
        </div>
      </div>
      
      <div class="chapter-navigation">
        <button 
          v-if="prevChapter" 
          @click="navigateToChapter(prevChapter)" 
          class="nav-button"
        >
          上一章
        </button>
        <button 
          v-if="nextChapter" 
          @click="navigateToChapter(nextChapter)" 
          class="nav-button"
        >
          下一章
        </button>
      </div>
    </div>
    
    <div v-else class="error">章节不存在</div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'

const route = useRoute()
const router = useRouter()

const novelId = Number(route.params.id)
const chapterId = Number(route.params.chapterId)

// 状态变量
const chapter = ref(null)
const chapters = ref([])
const novelTitle = ref('')
const loading = ref(true)
const hasAccess = ref(false)

// 阅读设置
const fontSize = ref(18)
const currentTheme = ref('light')

// 模拟数据
const mockChapters = {
  1: [
    {
      id: 1,
      novel_id: 1,
      title: '序章：地球末日',
      chapter_number: 0,
      is_paid: false
    },
    {
      id: 2,
      novel_id: 1,
      title: '起航',
      chapter_number: 1,
      is_paid: false
    },
    {
      id: 3,
      novel_id: 1,
      title: '深空探索',
      chapter_number: 2,
      is_paid: false
    },
    {
      id: 4,
      novel_id: 1,
      title: '第一次接触',
      chapter_number: 3,
      is_paid: true
    }
  ]
}

// 模拟章节内容
const mockChapterContents = {
  1: [
    '2150年，地球资源已经接近枯竭。',
    '人类在过去的百年间过度开采和消耗，使得地球的生态系统几乎崩溃。',
    '全球变暖导致海平面上升，许多沿海城市已经被淹没。',
    '极端天气事件频发，粮食生产受到严重影响。',
    '人类面临着前所未有的生存危机。',
    '在这种情况下，联合政府启动了「希望计划」，旨在寻找适合人类居住的新星球。',
    '经过多年的努力，科学家们终于找到了一个可能适合人类居住的星球——「新地球」。'
  ],
  2: [
    '2152年3月15日，这是一个值得人类铭记的日子。',
    '「希望号」星际飞船在全球数十亿人的注视下，缓缓离开了地球轨道。',
    '飞船上载着五百名经过严格选拔的宇航员和科学家，他们将作为人类的先锋，前往「新地球」。',
    '李明是飞船上的首席科学家，他望着窗外逐渐变小的地球，心中感慨万千。',
    '「再见了，地球。我们会找到新的家园的。」他默默地说道。',
    '飞船将以接近光速的速度飞行，预计需要五年时间才能到达目的地。'
  ],
  3: [
    '宇宙是如此的浩瀚和神秘。',
    '在过去的三年里，「希望号」穿越了无数的星系和星云。',
    '船员们见证了许多奇妙的宇宙现象：超新星爆发、黑洞吞噬恒星、星云的诞生...',
    '这些壮观的景象让他们忘记了旅途的枯燥和危险。',
    '然而，深空探索并非一帆风顺。',
    '他们遇到了强烈的宇宙射线暴，险些摧毁飞船的防护系统。',
    '还遭遇了陨石群，飞船的外壳受到了轻微的损伤。'
  ],
  4: [
    '2157年1月10日，「希望号」终于接近了「新地球」。',
    '当这个蓝色的星球出现在视野中时，整个飞船沸腾了。',
    '「我们成功了！」船员们欢呼雀跃。',
    '然而，就在他们准备进入轨道时，雷达突然检测到了异常信号。',
    '「有不明飞行物正在接近！」通信员惊呼道。',
    '李明立刻来到指挥室，看到屏幕上显示着几个快速移动的光点。',
    '「难道这个星球上已经有智慧生命了？」他自言自语道。'
  ]
}

// AI内容
const aiImages = ref([
  {
    id: 1,
    url: '/ai-image1.jpg',
    description: '宇宙飞船穿越星云的壮观景象'
  },
  {
    id: 2,
    url: '/ai-image2.jpg',
    description: '新地球的全景图'
  }
])

const aiVideos = ref([
  {
    id: 1,
    url: '/ai-video1.mp4',
    description: '宇宙飞船飞行模拟视频'
  }
])

// 章节内容段落
const chapterParagraphs = computed(() => {
  if (!chapter.value) return []
  return mockChapterContents[chapterId] || [
    '章节内容加载中...',
    '这是一个付费章节，您需要购买后才能阅读完整内容。'
  ]
})

// 计算前后章节
const prevChapter = computed(() => {
  const currentIndex = chapters.value.findIndex(c => c.id === chapterId)
  return currentIndex > 0 ? chapters.value[currentIndex - 1] : null
})

const nextChapter = computed(() => {
  const currentIndex = chapters.value.findIndex(c => c.id === chapterId)
  return currentIndex < chapters.value.length - 1 ? chapters.value[currentIndex + 1] : null
})

// 内容样式
const contentStyle = computed(() => ({
  fontSize: `${fontSize.value}px`,
  backgroundColor: getThemeColor('background'),
  color: getThemeColor('text')
}))

// 获取主题颜色
function getThemeColor() {
  const themes = {
    light: { background: '#fff', text: '#333' },
    dark: { background: '#333', text: '#eee' },
    sepia: { background: '#f4ecd8', text: '#5b4636' }
  }
  return themes[currentTheme.value] || themes.light
}

onMounted(async () => {
  try {
    // 模拟API请求延迟
    await new Promise(resolve => setTimeout(resolve, 500))
    
    // 获取章节列表
    chapters.value = mockChapters[novelId] || []
    
    // 获取当前章节
    chapter.value = chapters.value.find(c => c.id === chapterId)
    novelTitle.value = '星际迷航'
    
    // 检查用户权限
    // 这里模拟免费章节直接有权限
    if (!chapter.value?.is_paid) {
      hasAccess.value = true
    }
  } catch (error) {
    console.error('加载章节失败:', error)
  } finally {
    loading.value = false
  }
})

// 导航到其他章节
function navigateToChapter(targetChapter) {
  router.push(`/novels/${novelId}/chapters/${targetChapter.id}`)
}

// 字体大小控制
function increaseFontSize() {
  if (fontSize.value < 28) fontSize.value += 2
}

function decreaseFontSize() {
  if (fontSize.value > 12) fontSize.value -= 2
}

// 切换主题
function changeTheme() {
  // 主题切换逻辑
}

// 购买章节
function buyChapter() {
  // 模拟购买成功
  alert('购买成功！现在您可以阅读此章节了。')
  hasAccess.value = true
}

// 保存阅读进度
watch(() => chapter.value, (newChapter) => {
  if (newChapter) {
    // 保存阅读进度逻辑
    console.log(`保存进度: 小说${novelId}, 章节${chapterId}`)
  }
})
</script>

<style scoped>
.chapter-reading-container {
  max-width: 800px;
  margin: 0 auto;
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

.chapter-header {
  text-align: center;
  margin-bottom: 30px;
}

.chapter-header h1 {
  font-size: 28px;
  margin-bottom: 10px;
}

.chapter-info {
  color: #666;
  font-size: 16px;
}

.paid-badge {
  color: #ff6b6b;
  font-weight: bold;
}

.reading-settings {
  display: flex;
  justify-content: space-between;
  margin-bottom: 20px;
  padding: 10px 0;
  border-bottom: 1px solid #eee;
}

.font-size-control {
  display: flex;
  align-items: center;
  gap: 15px;
}

.setting-btn {
  background-color: #f0f0f0;
  border: none;
  padding: 5px 10px;
  border-radius: 4px;
  cursor: pointer;
}

.theme-switch select {
  padding: 5px 10px;
  border-radius: 4px;
}

.chapter-content {
  padding: 30px;
  border-radius: 8px;
  margin-bottom: 30px;
  min-height: 400px;
  transition: background-color 0.3s, color 0.3s;
}

.content-paragraph {
  margin-bottom: 1.5em;
  line-height: 1.8;
  text-indent: 2em;
}

.paid-content-blocked {
  text-align: center;
  padding: 50px;
  color: #666;
}

.buy-button {
  background-color: #ff6b6b;
  color: white;
  border: none;
  padding: 10px 20px;
  border-radius: 4px;
  font-size: 16px;
  cursor: pointer;
  margin-top: 20px;
}

.ai-content-section {
  margin-top: 40px;
  padding-top: 20px;
  border-top: 1px solid #eee;
}

.ai-content-section h3 {
  margin-bottom: 20px;
  color: #444;
}

.ai-images {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  gap: 20px;
}

.ai-image-item img {
  width: 100%;
  border-radius: 8px;
  margin-bottom: 10px;
}

.image-description {
  font-size: 14px;
  color: #666;
  text-align: center;
}

.ai-video-item {
  margin-bottom: 20px;
}

.ai-video-item video {
  width: 100%;
  border-radius: 8px;
}

.video-description {
  font-size: 14px;
  color: #666;
  text-align: center;
  margin-top: 10px;
}

.chapter-navigation {
  display: flex;
  justify-content: space-between;
  margin-top: 40px;
}

.nav-button {
  padding: 10px 20px;
  background-color: #2196F3;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 16px;
}

.nav-button:disabled {
  background-color: #ccc;
  cursor: not-allowed;
}

@media (max-width: 768px) {
  .reading-settings {
    flex-direction: column;
    gap: 10px;
  }
  
  .chapter-content {
    padding: 20px;
  }
  
  .ai-images {
    grid-template-columns: 1fr;
  }
}
</style>