export default defineNuxtConfig({
  devtools: {
    enabled: true
  },
  

  
  // 禁用telemetry以避免交互式提示
  telemetry: {
    enabled: false
  },
  
  // 使用系统建议的compatibilityDate
  compatibilityDate: '2025-11-27',
  
  modules: [
  ],
  css: [
    '@/assets/css/main.css'
  ],
  runtimeConfig: {
    public: {
      apiBase: '/api'
    },
    jwtSecret: process.env.JWT_SECRET || 'your-secret-key'
  },
  app: {
    head: {
      title: 'AI小说阅读平台',
      meta: [
        { name: 'description', content: '创新的AI生成小说阅读平台' },
        { name: 'viewport', content: 'width=device-width, initial-scale=1' }
      ],
      link: [
        { rel: 'icon', type: 'image/x-icon', href: '/favicon.ico' }
      ]
    }
  },
  router: {
    options: {
      scrollBehaviorType: 'smooth'
    },
    middleware: [] // 移除全局auth中间件以避免潜在冲突
  },
  
  // 移除可能导致问题的alias配置
  // alias配置已移除，使用默认的模块解析路径
  
  // 构建配置调整
  build: {
    // 移除h3的特殊转译，避免导入路径问题
    optimization: {
      minimize: false // 禁用压缩以便更好地调试
    }
  },
  
  // 配置Nitro服务器以解决h3兼容性问题
  nitro: {
    esbuild: {
      options: {
        target: 'es2022',
        keepNames: true
      }
    },
    runtimeConfig: {
      nodeOptions: '--no-warnings'
    },
    // 禁用服务器端的自动路由中间件生成
    autoImport: {
      dirs: [],
      components: false
    }
  }
})