<template>
  <div class="app-container">
    <header class="app-header">
      <h1 class="app-title">小说阅读平台</h1>
      <nav class="app-nav">
      <NuxtLink to="/" class="nav-link">首页</NuxtLink>
      <NuxtLink to="/novels" class="nav-link">小说列表</NuxtLink>
      <div class="navbar-user" v-if="!user">
        <NuxtLink to="/login" class="nav-link">登录</NuxtLink>
        <NuxtLink to="/register" class="nav-link btn-primary">注册</NuxtLink>
      </div>
      <div class="navbar-user" v-else>
        <span class="nav-link user-info">{{ user.username }}</span>
        <button @click="handleLogout" class="nav-link">登出</button>
      </div>
    </nav>
    </header>
    
    <main class="app-main">
      <NuxtPage />
    </main>
    
    <footer class="app-footer">
      <p>© 2024 小说阅读平台 - 创新阅读体验</p>
    </footer>
  </div>
</template>

<script setup lang="ts">
import { onMounted } from 'vue';

// 使用useState创建全局用户状态
const user = useState('user', () => null);

// 检查用户登录状态
const checkAuth = async () => {
  try {
    const response = await fetch('/api/auth/me', {
      credentials: 'include'
    });
    
    if (response.ok) {
      const result = await response.json();
      user.value = result.data.user;
    }
  } catch (error) {
    console.error('检查登录状态失败:', error);
  }
};

// 登出功能
const handleLogout = async () => {
  try {
    await fetch('/api/auth/logout', {
      method: 'POST',
      credentials: 'include'
    });
    user.value = null;
  } catch (error) {
    console.error('登出失败:', error);
  }
};

// 页面加载时检查登录状态
onMounted(() => {
  checkAuth();
});
</script>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  line-height: 1.6;
  color: #333;
  background-color: #f5f5f5;
}

.app-container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0 20px;
}

.app-header {
  background-color: #fff;
  padding: 20px 0;
  margin-bottom: 20px;
  border-bottom: 1px solid #eaeaea;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.app-title {
  font-size: 1.8rem;
  color: #333;
}

.app-nav {
  display: flex;
  gap: 20px;
}

.nav-link {
  color: #333;
  text-decoration: none;
  padding: 5px 10px;
  border-radius: 4px;
  transition: background-color 0.3s;
}

.nav-link:hover {
  background-color: #f0f0f0;
}

.btn-primary {
  background-color: #007bff;
  color: white;
  padding: 5px 15px;
  border-radius: 4px;
  text-decoration: none;
  transition: background-color 0.3s;
}

.btn-primary:hover {
  background-color: #0056b3 !important;
  color: white !important;
}

.navbar-user {
  display: flex;
  gap: 15px;
  align-items: center;
}

.user-info {
  font-weight: 500;
}

.navbar-user button {
  background: none;
  border: none;
  color: #333;
  cursor: pointer;
  padding: 5px 10px;
  border-radius: 4px;
  font-size: inherit;
}

.navbar-user button:hover {
  background-color: #f0f0f0;
}

.app-main {
  min-height: 600px;
  background-color: #fff;
  padding: 20px;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.app-footer {
  text-align: center;
  padding: 20px 0;
  margin-top: 20px;
  color: #666;
  font-size: 0.9rem;
}
</style>