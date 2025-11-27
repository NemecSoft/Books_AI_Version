// 最简单的API测试
export default (req, res) => {
  return {
    hello: 'world',
    timestamp: new Date().toISOString()
  }
};