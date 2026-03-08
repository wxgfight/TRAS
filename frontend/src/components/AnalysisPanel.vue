
<template>
  <div class="analysis-container" style="display: flex; gap: 20px; height: calc(100vh - 150px);">
    <!-- 左侧配置面板 -->
    <div class="config-panel" style="width: 300px; padding: 20px; background-color: white; border-radius: 8px; box-shadow: 0 2px 12px 0 rgba(0,0,0,0.1); overflow-y: auto;">
      <h3>配置图表</h3>
      <el-tabs v-model="activeTab" @tab-click="handleTabClick">
        <el-tab-pane label="新建分析" name="new">
          <el-form :model="form" label-position="top">
            <el-form-item label="选择文件">
              <el-select v-model="form.fileId" placeholder="请选择文件" @change="handleFileChange" style="width: 100%;">
                <el-option 
                  v-for="file in excelFiles" 
                  :key="file._id" 
                  :label="file.originalname" 
                  :value="file._id"
                ></el-option>
              </el-select>
            </el-form-item>
            
            <el-form-item label="选择Sheet">
              <el-select v-model="form.sheet" placeholder="请选择Sheet" @change="handleSheetChange" style="width: 100%;" :disabled="!form.fileId">
                <el-option v-for="sheet in sheets" :key="sheet" :label="sheet" :value="sheet"></el-option>
              </el-select>
            </el-form-item>
            
            <el-form-item label="图表类型">
              <el-select v-model="form.chartType" placeholder="选择图表类型" style="width: 100%;">
                <el-option label="柱状图" value="bar"></el-option>
                <el-option label="折线图" value="line"></el-option>
                <el-option label="饼图" value="pie"></el-option>
                <el-option label="散点图" value="scatter"></el-option>
              </el-select>
            </el-form-item>
            
            <el-form-item label="X轴 (维度)">
              <el-select v-model="form.xAxis" placeholder="选择X轴字段" style="width: 100%;" :disabled="!columns.length">
                <el-option 
                  v-for="col in columns" 
                  :key="col" 
                  :label="col.replace(/::/g, ' > ')" 
                  :value="col"
                />
              </el-select>
            </el-form-item>
            
            <el-form-item label="Y轴 (数值/度量)">
              <el-select v-model="form.yAxis" placeholder="选择Y轴字段" style="width: 100%;" :disabled="!columns.length">
                <el-option 
                  v-for="col in columns" 
                  :key="col" 
                  :label="col.replace(/::/g, ' > ')" 
                  :value="col"
                />
              </el-select>
            </el-form-item>
            
            <el-form-item label="聚合方式">
              <el-select v-model="form.aggregation" placeholder="选择聚合方式" style="width: 100%;">
                <el-option label="计数 (Count)" value="count"></el-option>
                <el-option label="求和 (Sum)" value="sum"></el-option>
                <el-option label="平均值 (Avg)" value="avg"></el-option>
                <el-option label="最大值 (Max)" value="max"></el-option>
                <el-option label="最小值 (Min)" value="min"></el-option>
              </el-select>
            </el-form-item>
            
            <div style="display: flex; gap: 10px;">
              <el-button type="primary" @click="runAnalysis" style="flex: 1;" :loading="loading" :disabled="!isReady">生成图表</el-button>
              <el-button type="success" @click="generateReport" style="flex: 1; padding-left: 5px; padding-right: 5px;" :loading="reportLoading" :disabled="!form.fileId || !form.sheet">生成统计报表</el-button>
            </div>
          </el-form>
        </el-tab-pane>
        <el-tab-pane label="历史记录" name="history">
          <div v-if="history.length > 0" class="history-list">
            <div v-for="item in history" :key="item._id" class="history-item" @click="loadHistory(item)">
              <div class="history-title">{{ item.result.title }}</div>
              <div class="history-meta">
                {{ item.fileId?.originalname || '未知文件' }}
                <br>
                {{ new Date(item.createdAt).toLocaleString() }}
              </div>
            </div>
          </div>
          <el-empty v-else description="暂无历史记录"></el-empty>
        </el-tab-pane>
      </el-tabs>
    </div>
    
    <!-- 右侧图表展示 -->
    <div class="chart-panel" style="flex: 1; background-color: white; border-radius: 8px; box-shadow: 0 2px 12px 0 rgba(0,0,0,0.1); padding: 20px; display: flex; flex-direction: column;">
      <div v-if="result" style="flex: 1; width: 100%; height: 100%; position: relative;">
        <!-- 使用 v-show 保持 DOM 存在，或者确保 v-if 切换时能正确获取 ref -->
        <div ref="chartRef" style="width: 100%; height: 100%; min-height: 400px;"></div>
      </div>
      <div v-else style="flex: 1; display: flex; justify-content: center; align-items: center; color: #909399;">
        <div style="text-align: center;">
          <el-icon :size="50"><PieChart /></el-icon>
          <p>请在左侧配置并生成图表</p>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, computed, watch, nextTick, onMounted, onBeforeUnmount } from 'vue'
import axios from 'axios'
import { PieChart } from '@element-plus/icons-vue'
// 确保 echarts 全局可用或按需引入，这里假设 window.echarts 或通过 import 引入
// 如果是 CDN 引入，window.echarts 可用。如果是 npm，需要 import
import * as echarts from 'echarts'

export default {
  name: 'AnalysisPanel',
  components: { PieChart },
  props: {
    excelFiles: {
      type: Array,
      default: () => []
    },
    token: {
      type: String,
      required: true
    }
  },
  setup(props) {
    const activeTab = ref('new')
    const form = ref({
      fileId: '',
      sheet: '',
      chartType: 'bar',
      xAxis: '',
      yAxis: '',
      aggregation: 'count',
      password: ''
    })
    
    const sheets = ref([])
    const columns = ref([])
    const result = ref(null)
    const history = ref([])
    const loading = ref(false)
    const reportLoading = ref(false)
    const chartRef = ref(null)
    let chartInstance = null
    let resizeObserver = null

    const isReady = computed(() => {
      return form.value.fileId && 
             form.value.sheet && 
             form.value.xAxis && 
             form.value.yAxis && 
             form.value.aggregation
    })

    // 获取历史记录
    const fetchHistory = async () => {
      try {
        const response = await axios.get('http://localhost:5000/api/analysis/history', {
          headers: { 'x-auth-token': props.token }
        })
        history.value = response.data
      } catch (error) {
        console.error('获取历史记录失败:', error)
      }
    }

    // 切换 Tab
    const handleTabClick = (tab) => {
      if (tab.paneName === 'history') {
        fetchHistory()
      }
    }

    // 文件变更
    const handleFileChange = async (fileId) => {
      form.value.sheet = ''
      form.value.xAxis = ''
      form.value.yAxis = ''
      sheets.value = []
      columns.value = []
      form.value.password = ''
      
      if (fileId) {
        await fetchColumns(fileId)
      }
    }

    // 获取列信息
    const fetchColumns = async (fileId, sheet = '') => {
      try {
        const params = {}
        if (sheet) params.sheet = sheet
        if (form.value.password) params.password = form.value.password
        
        const response = await axios.get(`http://localhost:5000/api/analysis/columns/${fileId}`, {
          headers: { 'x-auth-token': props.token },
          params
        })
        
        sheets.value = response.data.sheets
        columns.value = response.data.columns
        
        if (!form.value.sheet && response.data.sheet) {
          form.value.sheet = response.data.sheet
        }
      } catch (error) {
        console.error('获取列信息失败:', error)
        if (error.response?.data?.needPassword) {
           const password = prompt('该文件受密码保护，请输入密码:')
           if (password) {
             form.value.password = password
             await fetchColumns(fileId, sheet)
           } else {
             form.value.fileId = ''
           }
        } else {
           alert('获取文件信息失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      }
    }

    // Sheet 变更
    const handleSheetChange = async (sheet) => {
      if (form.value.fileId && sheet) {
        await fetchColumns(form.value.fileId, sheet)
      }
    }

    // 执行分析
    const runAnalysis = async () => {
      if (!isReady.value) return
      
      loading.value = true
      try {
        const response = await axios.post('http://localhost:5000/api/analysis/analyze', form.value, {
          headers: { 'x-auth-token': props.token }
        })
        
        result.value = response.data
        // 渲染图表
        renderChart()
        
        // 刷新历史
        fetchHistory()
        
      } catch (error) {
        console.error('分析失败:', error)
        if (error.response?.data?.needPassword) {
           const password = prompt('该文件受密码保护，请输入密码:')
           if (password) {
             form.value.password = password
             runAnalysis() // 重试
           }
        } else {
           alert('分析失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      } finally {
        loading.value = false
      }
    }

    // 生成综合报表
    const generateReport = async () => {
      if (!form.value.fileId || !form.value.sheet) return
      
      reportLoading.value = true
      try {
        const response = await axios.post('http://localhost:5000/api/analysis/report', {
          fileId: form.value.fileId,
          sheet: form.value.sheet,
          password: form.value.password
        }, {
          headers: { 'x-auth-token': props.token },
          responseType: 'blob' // 重要：接收二进制流
        })
        
        // 创建下载链接
        const url = window.URL.createObjectURL(new Blob([response.data]))
        const link = document.createElement('a')
        link.href = url
        
        // 获取文件名
        const contentDisposition = response.headers['content-disposition']
        let fileName = 'analysis_report.xlsx'
        if (contentDisposition) {
           const fileNameMatch = contentDisposition.match(/filename="?([^"]+)"?/)
           if (fileNameMatch && fileNameMatch.length === 2) fileName = fileNameMatch[1]
        }
        
        link.setAttribute('download', fileName)
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
        window.URL.revokeObjectURL(url)
        
      } catch (error) {
        console.error('生成报表失败:', error)
        if (error.response?.data && error.response.data instanceof Blob && error.response.data.type === 'application/json') {
           // 如果是JSON错误（如密码错误），读取blob
           const reader = new FileReader()
           reader.onload = () => {
             const errorData = JSON.parse(reader.result)
             if (errorData.needPassword) {
               const password = prompt('该文件受密码保护，请输入密码:')
               if (password) {
                 form.value.password = password
                 generateReport() // 重试
               }
             } else {
               alert('生成报表失败: ' + (errorData.msg || '服务器错误'))
             }
           }
           reader.readAsText(error.response.data)
        } else {
           // alert('生成报表失败，请稍后重试')
           console.log('blob error', error)
        }
      } finally {
        reportLoading.value = false
      }
    }

    // 加载历史记录
    const loadHistory = (item) => {
      result.value = item.result
      form.value = {
        fileId: item.fileId?._id || '',
        sheet: item.sheet,
        chartType: item.chartType,
        xAxis: item.xAxis,
        yAxis: item.yAxis,
        aggregation: item.aggregation,
        password: ''
      }
      renderChart()
    }

    // 渲染图表核心逻辑
    const renderChart = async () => {
      await nextTick()
      
      if (!chartRef.value) {
        // 如果 DOM 还没出来，重试一次
        setTimeout(renderChart, 100)
        return
      }

      if (chartInstance) {
        chartInstance.dispose()
      }

      chartInstance = echarts.init(chartRef.value)
      
      const { xAxis, series, chartType, title } = result.value
      
      // 美化配色方案
      const colors = ['#5470c6', '#91cc75', '#fac858', '#ee6666', '#73c0de', '#3ba272', '#fc8452', '#9a60b4', '#ea7ccc']

      const option = {
        color: colors,
        title: {
          text: title,
          left: 'center',
          textStyle: {
            fontSize: 20,
            fontWeight: 'bold',
            color: '#333'
          },
          subtextStyle: {
            color: '#aaa'
          }
        },
        tooltip: {
          trigger: chartType === 'pie' ? 'item' : 'axis',
          axisPointer: { type: 'shadow' },
          backgroundColor: 'rgba(255, 255, 255, 0.9)',
          borderColor: '#ccc',
          borderWidth: 1,
          textStyle: {
            color: '#333'
          }
        },
        legend: {
          bottom: 0,
          left: 'center',
          data: chartType === 'pie' ? xAxis : undefined,
          padding: [10, 0]
        },
        toolbox: {
          feature: {
            saveAsImage: { title: '保存图片' },
            dataView: { title: '数据视图' },
            magicType: { 
              type: ['line', 'bar'],
              title: { line: '切换为折线图', bar: '切换为柱状图' }
            },
            restore: { title: '还原' }
          },
          right: 20
        },
        grid: {
          left: '3%',
          right: '4%',
          bottom: '15%', // 增加底部空间给 legend 和 axisLabel
          containLabel: true,
          show: false // 不显示网格边框
        },
        xAxis: chartType === 'pie' ? undefined : {
          type: 'category',
          data: xAxis,
          axisLabel: {
            rotate: 45,
            interval: 0,
            color: '#666'
          },
          axisLine: {
            lineStyle: {
              color: '#ccc'
            }
          },
          axisTick: {
            alignWithLabel: true
          }
        },
        yAxis: chartType === 'pie' ? undefined : {
          type: 'value',
          splitLine: {
            lineStyle: {
              type: 'dashed',
              color: '#eee'
            }
          },
          axisLabel: {
            color: '#666'
          }
        },
        series: [{
          name: form.value.yAxis || '数值', // 给系列加个名字，tooltip 显示更好看
          data: chartType === 'pie' 
            ? xAxis.map((key, i) => ({ name: key, value: series[i] }))
            : series,
          type: chartType === 'scatter' ? 'scatter' : (chartType === 'pie' ? 'pie' : chartType),
          
          // 针对不同图表的特定美化
          ...(chartType === 'bar' ? {
            itemStyle: {
              borderRadius: [5, 5, 0, 0], // 圆角柱子
              // 渐变色
              color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                { offset: 0, color: '#83bff6' },
                { offset: 0.5, color: '#188df0' },
                { offset: 1, color: '#188df0' }
              ])
            },
            emphasis: {
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#2378f7' },
                  { offset: 0.7, color: '#2378f7' },
                  { offset: 1, color: '#83bff6' }
                ])
              }
            },
            barMaxWidth: 50
          } : {}),

          ...(chartType === 'line' ? {
            smooth: true, // 平滑曲线
            symbol: 'circle',
            symbolSize: 8,
            lineStyle: {
              width: 3,
              shadowColor: 'rgba(0,0,0,0.3)',
              shadowBlur: 10,
              shadowOffsetY: 8
            },
            areaStyle: {
              opacity: 0.3,
              color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                { offset: 0, color: 'rgba(128, 255, 165, 0.8)' },
                { offset: 1, color: 'rgba(1, 191, 236, 0.1)' }
              ])
            }
          } : {}),

          ...(chartType === 'pie' ? {
            radius: ['40%', '70%'], // 环形图
            avoidLabelOverlap: false,
            itemStyle: {
              borderRadius: 10,
              borderColor: '#fff',
              borderWidth: 2
            },
            label: {
              show: false,
              position: 'center'
            },
            emphasis: {
              label: {
                show: true,
                fontSize: '20',
                fontWeight: 'bold'
              },
              itemStyle: {
                shadowBlur: 10,
                shadowOffsetX: 0,
                shadowColor: 'rgba(0, 0, 0, 0.5)'
              }
            },
            labelLine: {
              show: false
            }
          } : {}),
          
          ...(chartType === 'scatter' ? {
             symbolSize: 12,
             itemStyle: {
                shadowBlur: 10,
                shadowColor: 'rgba(120, 36, 50, 0.5)',
                shadowOffsetY: 5,
                color: new echarts.graphic.RadialGradient(0.4, 0.3, 1, [{
                    offset: 0,
                    color: 'rgb(251, 118, 123)'
                }, {
                    offset: 1,
                    color: 'rgb(204, 46, 72)'
                }])
            }
          } : {})
        }]
      }
      
      chartInstance.setOption(option)
    }

    // 监听窗口大小变化 (Robust implementation)
    onMounted(() => {
      fetchHistory()
      
      resizeObserver = new ResizeObserver(() => {
        if (chartInstance) {
          chartInstance.resize()
        }
      })
      // 这里的 chartRef 可能因为 v-if="result" 还是 null
      // 我们需要监听 chartRef 的变化，或者在 renderChart 里绑定 observer
    })

    // 监听 result 变化来绑定 ResizeObserver
    watch(result, async (val) => {
      if (val) {
        await nextTick()
        if (chartRef.value && resizeObserver) {
          resizeObserver.observe(chartRef.value)
        }
      } else {
        if (chartInstance) {
          chartInstance.dispose()
          chartInstance = null
        }
      }
    })

    onBeforeUnmount(() => {
      if (chartInstance) {
        chartInstance.dispose()
      }
      if (resizeObserver) {
        resizeObserver.disconnect()
      }
    })

    return {
      activeTab,
      form,
      sheets,
      columns,
      result,
      history,
      loading,
      reportLoading,
      chartRef,
      isReady,
      handleTabClick,
      handleFileChange,
      handleSheetChange,
      runAnalysis,
      generateReport,
      loadHistory
    }
  }
}
</script>

<style scoped>
.history-item {
  padding: 10px;
  border-bottom: 1px solid #eee;
  cursor: pointer;
  transition: background-color 0.3s;
}
.history-item:hover {
  background-color: #f5f7fa;
}
.history-title {
  font-weight: bold;
  font-size: 14px;
  margin-bottom: 5px;
}
.history-meta {
  font-size: 12px;
  color: #909399;
}
</style>
