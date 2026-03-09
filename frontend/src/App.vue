<template>
  <div id="app">
    <el-container v-if="isLoggedIn">
      <el-header>
        <div class="header-content">
          <h1>测试报告分析系统</h1>
          <div class="user-info">
            <span>{{ user.username }}</span>
            <el-button type="primary" @click="logout">退出登录</el-button>
          </div>
        </div>
      </el-header>
      <el-container>
        <el-aside width="200px">
          <el-menu
            :default-active="activeMenu"
            class="el-menu-vertical-demo"
            @select="handleMenuSelect"
          >
            <el-menu-item index="upload">
              <el-icon><upload-filled /></el-icon>
              <span>数据上传</span>
            </el-menu-item>
            <el-menu-item index="data">
              <el-icon><document /></el-icon>
              <span>数据管理</span>
            </el-menu-item>
            <el-menu-item index="visualization">
              <el-icon><pie-chart /></el-icon>
              <span>数据可视化</span>
            </el-menu-item>
            <el-menu-item index="compare">
              <el-icon><document /></el-icon>
              <span>文件对比</span>
            </el-menu-item>
            <el-menu-item index="history">
              <el-icon><document /></el-icon>
              <span>历史记录</span>
            </el-menu-item>
            <el-menu-item index="dashboard">
              <el-icon><pie-chart /></el-icon>
              <span>数据看板</span>
            </el-menu-item>
          </el-menu>
        </el-aside>
        <el-main>
          <div v-if="activeMenu === 'upload'">
            <el-tabs v-model="activeUploadTab" type="border-card">
              <el-tab-pane label="数据上传" name="dataUpload">
                <el-upload
                  ref="uploadRef"
                  class="upload-demo"
                  action="http://localhost:5000/api/data/upload"
                  :headers="{ 'x-auth-token': token }"
                  :data="{ headerRowStart: uploadHeaderRowStart, headerRowEnd: uploadHeaderRowEnd, password: uploadPassword }"
                  :on-success="handleUploadSuccess"
                  :on-error="handleUploadError"
                  :auto-upload="false"
                  :file-list="fileList"
                  :limit="1"
                  accept=".csv,.xlsx,.xls,.json,.txt"
                >
                  <el-button type="primary">选择文件</el-button>
                  <template #tip>
                    <div class="el-upload__tip">
                      支持上传 CSV、Excel、JSON、TXT 文件，最大 50MB
                    </div>
                  </template>
                </el-upload>
                <div style="margin-top: 15px; margin-bottom: 15px;">
                  <span style="margin-right: 10px;">表头行范围:</span>
                  <el-input-number v-model="uploadHeaderRowStart" :min="0" :max="10" placeholder="起始行" controls-position="right" style="width: 100px;"></el-input-number>
                  <span style="margin: 0 10px;">至</span>
                  <el-input-number v-model="uploadHeaderRowEnd" :min="0" :max="10" placeholder="结束行" controls-position="right" style="width: 100px;"></el-input-number>
                  <span style="margin-left: 10px; color: #909399; font-size: 12px;">(0表示第一行)</span>
                </div>
                <div style="margin-top: 15px; margin-bottom: 15px;">
                  <span style="margin-right: 10px;">Excel密码:</span>
                  <el-input v-model="uploadPassword" placeholder="若文件加密请填写" style="width: 200px;" show-password></el-input>
                  <span style="margin-left: 10px; color: #909399; font-size: 12px;">(可选)</span>
                </div>
                <el-button type="success" @click="submitUpload">
                  开始上传
                </el-button>
              </el-tab-pane>
              <el-tab-pane label="模板管理" name="templateManagement">
                <div class="template-actions" style="margin-bottom: 20px;">
                  <el-button type="primary" @click="templateUploadDialogVisible = true">上传模板</el-button>
                </div>
                <el-table :data="templateList" style="width: 100%">
                  <el-table-column type="index" label="序号" width="80"></el-table-column>
                  <el-table-column prop="originalname" label="模板名称"></el-table-column>
                  <el-table-column prop="createdAt" label="更新时间" :formatter="formatDate"></el-table-column>
                  <el-table-column label="操作" width="250">
                    <template #default="scope">
                      <el-button type="primary" size="small" @click="previewTemplate(scope.row)" v-if="scope.row.originalname.endsWith('.xlsx') || scope.row.originalname.endsWith('.xls')">预览</el-button>
                      <el-button type="success" size="small" @click="downloadTemplate(scope.row)">下载</el-button>
                      <el-button type="danger" size="small" @click="deleteTemplate(scope.row._id)">删除</el-button>
                    </template>
                  </el-table-column>
                </el-table>
              </el-tab-pane>
            </el-tabs>

            <!-- Template Upload Dialog -->
            <el-dialog v-model="templateUploadDialogVisible" title="上传模板" width="30%">
              <el-upload
                ref="templateUploadRef"
                class="upload-demo"
                action="http://localhost:5000/api/template/upload"
                :headers="{ 'x-auth-token': token }"
                :on-success="handleTemplateUploadSuccess"
                :on-error="handleTemplateUploadError"
                :auto-upload="false"
                :file-list="templateFileList"
                :limit="1"
                accept=".xlsx,.xls"
              >
                <template #trigger>
                  <el-button type="primary">选择文件</el-button>
                </template>
                <el-button style="margin-left: 10px;" type="success" @click="submitTemplateUpload">上传到服务器</el-button>
                <template #tip>
                  <div class="el-upload__tip">
                    只能上传 xlsx/xls 文件，且不超过 50MB
                  </div>
                </template>
              </el-upload>
            </el-dialog>
          </div>
          <div v-else-if="activeMenu === 'data'">
            <el-table :data="dataList" style="width: 100%">
              <el-table-column prop="projectName" label="项目名称" width="180"></el-table-column>
              <el-table-column prop="reportName" label="报告名称" width="220"></el-table-column>
              <el-table-column prop="originalname" label="文件名" min-width="300"></el-table-column>
              <el-table-column label="文件类型" width="100">
                <template #default="scope">
                  {{ scope.row.originalname.split('.').pop().toUpperCase() }}
                </template>
              </el-table-column>
              <el-table-column prop="size" label="文件大小" :formatter="formatSize"></el-table-column>
              <el-table-column prop="createdAt" label="上传时间" :formatter="formatDate"></el-table-column>
              <el-table-column label="加密" width="80" align="center">
                <template #default="scope">
                  <el-tag type="danger" size="small" v-if="isEncrypted(scope.row)">是</el-tag>
                  <el-tag type="info" size="small" v-else>否</el-tag>
                </template>
              </el-table-column>
              <el-table-column label="操作" width="250">
                <template #default="scope">
                  <el-button type="primary" size="small" @click="previewFile(scope.row)" v-if="scope.row.originalname.endsWith('.xlsx') || scope.row.originalname.endsWith('.xls')">预览</el-button>
                  <el-button type="success" size="small" @click="downloadFile(scope.row)">下载</el-button>
                  <el-button type="danger" size="small" @click="deleteData(scope.row._id)">
                    删除
                  </el-button>
                </template>
              </el-table-column>
            </el-table>
            

          </div>
          <div v-else-if="activeMenu === 'visualization'">
            <AnalysisPanel :excel-files="excelFiles" :token="token" />
          </div>
          <div v-else-if="activeMenu === 'compare'">
            <el-form :model="compareForm" label-width="100px">
              <el-form-item label="第一个文件">
                <el-select v-model="compareForm.file1Id" @change="handleFile1Change" :placeholder="excelFiles.length > 0 ? '请选择文件' : '暂无Excel文件，请先上传'">
                  <el-option 
                    v-for="file in excelFiles" 
                    :key="file._id" 
                    :label="file.originalname" 
                    :value="file._id"
                  ></el-option>
                </el-select>
              </el-form-item>
              <el-form-item label="第一个文件Sheet">
                <el-select v-model="compareForm.sheet1" @change="handleSheet1Change">
                  <el-option 
                    v-for="sheet in sheets1" 
                    :key="sheet" 
                    :label="sheet" 
                    :value="sheet"
                  ></el-option>
                </el-select>
              </el-form-item>
              <el-form-item label="第二个文件">
                <el-select v-model="compareForm.file2Id" @change="handleFile2Change" :placeholder="excelFiles.length > 0 ? '请选择文件' : '暂无Excel文件，请先上传'">
                  <el-option 
                    v-for="file in excelFiles" 
                    :key="file._id" 
                    :label="file.originalname" 
                    :value="file._id"
                  ></el-option>
                </el-select>
              </el-form-item>
              <el-form-item label="第二个文件Sheet">
                <el-select v-model="compareForm.sheet2">
                  <el-option 
                    v-for="sheet in sheets2" 
                    :key="sheet" 
                    :label="sheet" 
                    :value="sheet"
                  ></el-option>
                </el-select>
              </el-form-item>
              <!-- 主键列已默认全选并隐藏
              <el-form-item label="主键列">
                <el-select v-model="compareForm.primaryKeys" multiple placeholder="选择主键列">
                  <el-option 
                    v-for="column in columns1" 
                    :key="column" 
                    :label="column" 
                    :value="column"
                  ></el-option>
                </el-select>
                <div style="margin-top: 5px; font-size: 12px; color: #606266;">
                  支持多列组合主键
                </div>
              </el-form-item>
              -->
              <el-form-item label="忽略列">
                <el-select v-model="compareForm.ignoreColumns" multiple placeholder="选择忽略列">
                  <el-option 
                    v-for="column in columns1" 
                    :key="column" 
                    :label="column" 
                    :value="column"
                  ></el-option>
                </el-select>
              </el-form-item>
              <el-form-item label="表头行范围">
                <el-input-number v-model="compareForm.headerRowStart" :min="0" :max="10" placeholder="起始" controls-position="right" style="width: 100px;"></el-input-number>
                <span style="margin: 0 10px;">至</span>
                <el-input-number v-model="compareForm.headerRowEnd" :min="0" :max="10" placeholder="结束" controls-position="right" style="width: 100px;"></el-input-number>
              </el-form-item>
              <el-form-item label="空值处理策略">
                <el-select v-model="compareForm.nullValueStrategy">
                  <el-option label="忽略" value="ignore"></el-option>
                  <el-option label="视为空字符串" value="treat_as_empty"></el-option>
                  <el-option label="视为0" value="treat_as_zero"></el-option>
                </el-select>
              </el-form-item>
              <el-form-item label="检测移动记录">
                <el-switch v-model="compareForm.detectMoved"></el-switch>
                <div style="margin-top: 5px; font-size: 12px; color: #606266;">
                  检测位置变化但内容不变的记录
                </div>
              </el-form-item>
              <el-form-item>
                <el-button type="primary" @click="compareFiles">开始对比</el-button>
              </el-form-item>
            </el-form>
            <div v-if="comparisonResult" style="margin-top: 20px;">
              <!-- 对比详情已隐藏
              <h3>对比结果</h3>
              <p>文件1: {{ comparisonResult.file1 }}</p>
              <p>文件2: {{ comparisonResult.file2 }}</p>
              <p>Sheet1: {{ comparisonResult.sheet1 }}</p>
              <p>Sheet2: {{ comparisonResult.sheet2 }}</p>
              <p>使用的主键: {{ comparisonResult.primaryKeys.join(', ') }}</p>
              <p>推荐的主键: {{ comparisonResult.recommendedPrimaryKeys.join(', ') }}</p>
              <p>忽略的列: {{ comparisonResult.ignoreColumns.join(', ') || '无' }}</p>
              <p>空值处理策略: {{ comparisonResult.nullValueStrategy }}</p>
              <p>检测移动记录: {{ comparisonResult.detectMoved ? '是' : '否' }}</p>
              -->
              
              <div style="margin-top: 20px; text-align: center;">
                 <el-button type="success" size="large" icon="Download" @click="downloadReport(comparisonResult.taskId)">下载对比报告</el-button>
              </div>

              <!-- 差异详情已隐藏，仅提供下载
              <h4 style="margin-top: 20px;">差异详情</h4>
              <el-table :data="comparisonResult.differences" style="width: 100%; margin-top: 10px;">
                <el-table-column prop="type" label="差异类型" :formatter="formatDifferenceType"></el-table-column>
                <el-table-column prop="file" label="文件"></el-table-column>
                <el-table-column label="行索引">
                  <template #default="scope">
                    <span v-if="scope.row.type === 'modified'">
                      {{ scope.row.rowIndex1 }} → {{ scope.row.rowIndex2 }}
                    </span>
                    <span v-else-if="scope.row.type === 'moved'">
                      {{ scope.row.oldIndex }} → {{ scope.row.newIndex }}
                    </span>
                    <span v-else>
                      {{ scope.row.rowIndex }}
                    </span>
                  </template>
                </el-table-column>
                <el-table-column prop="column" label="列"></el-table-column>
                <el-table-column prop="oldValue" label="旧值"></el-table-column>
                <el-table-column prop="newValue" label="新值"></el-table-column>
                <el-table-column label="变更详情">
                  <template #default="scope">
                    <div v-if="scope.row.type === 'modified' && scope.row.changes">
                      <div v-for="(change, index) in scope.row.changes" :key="index">
                        {{ change.column }}: {{ change.oldValue }} → {{ change.newValue }}
                      </div>
                    </div>
                    <div v-else>
                      {{ scope.row.rowData ? JSON.stringify(scope.row.rowData) : '' }}
                    </div>
                  </template>
                </el-table-column>
                <el-table-column prop="primaryKey" label="主键值"></el-table-column>
              </el-table>
              <el-button type="success" style="margin-top: 10px;" @click="downloadReport(comparisonResult.taskId)">下载对比报告</el-button>
              -->
            </div>
          </div>
          <div v-else-if="activeMenu === 'history'">
            <el-form :inline="true" :model="historyForm" style="margin-bottom: 20px;">
              <el-form-item label="搜索">
                <el-input v-model="historyForm.search" placeholder="文件名或任务ID"></el-input>
              </el-form-item>
              <el-form-item label="时间范围">
                <el-date-picker v-model="historyForm.dateRange" type="daterange" range-separator="至" start-placeholder="开始日期" end-placeholder="结束日期"></el-date-picker>
              </el-form-item>
              <el-form-item>
                <el-button type="primary" @click="fetchHistoryList">查询</el-button>
                <el-button @click="resetHistoryForm">重置</el-button>
              </el-form-item>
            </el-form>
            <el-table :data="historyList" style="width: 100%;">
              <el-table-column prop="taskId" label="任务ID" width="180"></el-table-column>
              <el-table-column label="文件1">
                <template #default="scope">
                  {{ scope.row.file1Id ? scope.row.file1Id.originalname : '文件已删除' }}
                </template>
              </el-table-column>
              <el-table-column label="文件2">
                <template #default="scope">
                  {{ scope.row.file2Id ? scope.row.file2Id.originalname : '文件已删除' }}
                </template>
              </el-table-column>
              <el-table-column prop="createdAt" label="对比时间" :formatter="formatDate" width="160"></el-table-column>
              <el-table-column label="差异统计" width="200">
                <template #default="scope">
                  <div style="font-size: 12px;">
                    <span style="color: #67C23A">新增: {{ scope.row.stats?.added || 0 }}</span> |
                    <span style="color: #F56C6C">删除: {{ scope.row.stats?.deleted || 0 }}</span> |
                    <span style="color: #E6A23C">修改: {{ scope.row.stats?.modified || 0 }}</span>
                  </div>
                </template>
              </el-table-column>
              <el-table-column prop="status" label="状态" width="80">
                 <template #default="scope">
                   <el-tag :type="scope.row.status === 'completed' ? 'success' : 'info'">{{ scope.row.status === 'completed' ? '完成' : '进行中' }}</el-tag>
                 </template>
              </el-table-column>
              <el-table-column label="操作" width="250">
                <template #default="scope">
                  <el-button type="primary" size="small" @click="downloadReport(scope.row.taskId)" v-if="scope.row.status === 'completed'">下载报表</el-button>
                </template>
              </el-table-column>
            </el-table>
            <el-pagination
              v-model:current-page="historyPage"
              v-model:page-size="historyLimit"
              :page-sizes="[10, 20, 50, 100]"
              layout="total, sizes, prev, pager, next, jumper"
              :total="historyTotal"
              @size-change="handleSizeChange"
              @current-change="handleCurrentChange"
              style="margin-top: 20px;"
            ></el-pagination>
          </div>
          <div v-else-if="activeMenu === 'dashboard'">
            <div class="dashboard-cards" style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 20px;">
              <el-card shadow="hover">
                <template #header>
                  <div class="card-header">
                    <span>总文件数</span>
                  </div>
                </template>
                <div class="card-content">
                  <el-statistic :value="dashboardStats.totalFiles" :precision="0"></el-statistic>
                </div>
              </el-card>
              <el-card shadow="hover">
                <template #header>
                  <div class="card-header">
                    <span>总对比任务数</span>
                  </div>
                </template>
                <div class="card-content">
                  <el-statistic :value="dashboardStats.totalComparisons" :precision="0"></el-statistic>
                </div>
              </el-card>
              <el-card shadow="hover">
                <template #header>
                  <div class="card-header">
                    <span>今日任务数</span>
                  </div>
                </template>
                <div class="card-content">
                  <el-statistic :value="dashboardStats.todayComparisons" :precision="0"></el-statistic>
                </div>
              </el-card>
              <el-card shadow="hover">
                <template #header>
                  <div class="card-header">
                    <span>近7天任务趋势</span>
                  </div>
                </template>
                <div class="card-content">
                  <div ref="trendChart" style="width: 100%; height: 150px;"></div>
                </div>
              </el-card>
            </div>
            <div class="dashboard-charts" style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
              <el-card shadow="hover">
                <template #header>
                  <div class="card-header">
                    <span>差异类型分布</span>
                  </div>
                </template>
                <div class="card-content">
                  <div ref="distributionChart" style="width: 100%; height: 300px;"></div>
                </div>
              </el-card>
              <el-card shadow="hover">
                <template #header>
                  <div class="card-header">
                    <span>任务时间趋势</span>
                  </div>
                </template>
                <div class="card-content">
                  <div ref="timeChart" style="width: 100%; height: 300px;"></div>
                </div>
              </el-card>
            </div>
          </div>
            <!-- 预览对话框 -->
            <el-dialog 
              v-model="previewDialogVisible" 
              title="文件预览" 
              width="90%" 
              top="5vh"
              class="preview-dialog"
              destroy-on-close
            >
              <div v-if="previewData" class="preview-content">
                <!-- 文件元数据 -->
                <el-descriptions :column="4" border size="small" class="mb-4">
                  <el-descriptions-item label="文件名">{{ previewData.filename }}</el-descriptions-item>
                  <el-descriptions-item label="Sheet数量">{{ previewData.sheetCount || 1 }}</el-descriptions-item>
                  <el-descriptions-item label="总行数">{{ previewData.totalRows }}</el-descriptions-item>
                  <el-descriptions-item label="当前Sheet">{{ previewParams.sheet || '默认' }}</el-descriptions-item>
                </el-descriptions>

                <!-- 工具栏 -->
                <div class="preview-toolbar">
                  <div class="toolbar-left">
                    <div class="toolbar-item" v-if="previewData.sheets && previewData.sheets.length > 0">
                      <span class="label">Sheet:</span>
                      <el-select v-model="previewParams.sheet" placeholder="选择Sheet" @change="handlePreviewParamsChange" size="default" style="width: 180px;">
                        <el-option v-for="sheet in previewData.sheets" :key="sheet" :label="sheet" :value="sheet"></el-option>
                      </el-select>
                    </div>

                    <div class="toolbar-item">
                      <span class="label">表头范围:</span>
                      <el-input-number v-model="previewParams.headerRowStart" :min="0" :max="10" size="default" @change="handlePreviewParamsChange" style="width: 100px;"></el-input-number>
                      <span style="margin: 0 5px;">-</span>
                      <el-input-number v-model="previewParams.headerRowEnd" :min="0" :max="10" size="default" @change="handlePreviewParamsChange" style="width: 100px;"></el-input-number>
                    </div>
                    
                    <div class="toolbar-item">
                      <span class="label">显示行数:</span>
                      <el-select v-model="previewParams.limit" placeholder="行数" @change="handlePreviewParamsChange" size="default" style="width: 100px;">
                        <el-option label="30行" :value="30"></el-option>
                        <el-option label="50行" :value="50"></el-option>
                        <el-option label="100行" :value="100"></el-option>
                        <el-option label="200行" :value="200"></el-option>
                      </el-select>
                    </div>
                  </div>

                  <div class="toolbar-right">
                    <div class="toolbar-item" v-if="previewNeedPassword">
                      <span class="label error-text">需密码:</span>
                      <el-input v-model="previewParams.password" type="password" placeholder="输入密码" size="default" style="width: 150px;" @keyup.enter="handlePreviewParamsChange" show-password></el-input>
                      <el-button type="primary" size="default" @click="handlePreviewParamsChange">解锁</el-button>
                    </div>
                    <el-button type="primary" plain icon="Refresh" @click="handlePreviewParamsChange">刷新预览</el-button>
                  </div>
                </div>

                <!-- 数据表格 -->
                <el-table 
                  :data="previewData.preview" 
                  height="60vh" 
                  border 
                  stripe
                  highlight-current-row
                  v-loading="previewLoading"
                  :header-cell-style="{ background: '#f5f7fa', color: '#606266', fontWeight: 'bold' }"
                >
                  <el-table-column type="index" label="#" width="60" fixed></el-table-column>
                  <el-table-column 
                    v-for="(value, key) in (previewData.preview[0] || {})" 
                    :key="key" 
                    :prop="key" 
                    :label="key"
                    min-width="150"
                    show-overflow-tooltip
                  >
                    <template #header>
                      <div style="white-space: pre-wrap; line-height: 1.2;">{{ key.replace(/::/g, '\n') }}</div>
                    </template>
                  </el-table-column>
                </el-table>
              </div>
              <div v-else v-loading="previewLoading" class="empty-preview">
                <el-empty description="暂无预览数据" v-if="!previewLoading"></el-empty>
              </div>
            </el-dialog>
        </el-main>
      </el-container>
    </el-container>
    <div v-else class="login-container">
      <div class="login-box">
        <div class="login-header">
          <img src="https://element-plus.org/images/element-plus-logo.svg" alt="logo" class="login-logo">
          <h2>TRAS</h2>
          <p>测试报告分析系统</p>
        </div>
        <el-tabs v-model="activeTab" stretch>
          <el-tab-pane label="登录" name="login">
            <el-form :model="loginForm" label-position="top" size="large">
              <el-form-item label="用户名">
                <el-input v-model="loginForm.username" placeholder="请输入用户名" prefix-icon="User"></el-input>
              </el-form-item>
              <el-form-item label="密码">
                <el-input v-model="loginForm.password" type="password" placeholder="请输入密码" prefix-icon="Lock" show-password @keyup.enter="login"></el-input>
              </el-form-item>
              <el-form-item>
                <el-button type="primary" class="login-btn" @click="login" :loading="loginLoading">登录</el-button>
              </el-form-item>
            </el-form>
          </el-tab-pane>
          <el-tab-pane label="注册" name="register">
            <el-form :model="registerForm" label-position="top" size="large">
              <el-form-item label="用户名">
                <el-input v-model="registerForm.username" placeholder="请输入用户名" prefix-icon="User"></el-input>
              </el-form-item>
              <el-form-item label="邮箱">
                <el-input v-model="registerForm.email" placeholder="请输入邮箱" prefix-icon="Message"></el-input>
              </el-form-item>
              <el-form-item label="密码">
                <el-input v-model="registerForm.password" type="password" placeholder="请输入密码" prefix-icon="Lock" show-password></el-input>
              </el-form-item>
              <el-form-item>
                <el-button type="success" class="login-btn" @click="register" :loading="registerLoading">注册</el-button>
              </el-form-item>
            </el-form>
          </el-tab-pane>
        </el-tabs>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, onMounted, computed, nextTick, watch } from 'vue'
import axios from 'axios'
import { UploadFilled, Document, PieChart, Printer, Refresh, User, Lock, Message } from '@element-plus/icons-vue'
import AnalysisPanel from './components/AnalysisPanel.vue'

export default {
  name: 'App',
  components: {
    UploadFilled,
    Document,
    PieChart,
    Printer,
    Refresh,
    User,
    Lock,
    Message,
    AnalysisPanel
  },
  setup() {
    const isLoggedIn = ref(false)
    const user = ref({})
    const token = ref('')
    const activeMenu = ref('upload')
    const activeUploadTab = ref('dataUpload')
    const activeTab = ref('login')
    const fileList = ref([])
    const dataList = ref([])
    const templateList = ref([])
    const templateUploadDialogVisible = ref(false)
    const templateFileList = ref([])
    const templateUploadRef = ref(null)
    const isPreviewingTemplate = ref(false)
    const loginLoading = ref(false)
    const registerLoading = ref(false)
    
    const excelFiles = computed(() => {
      return dataList.value.filter(file => 
        file.originalname.endsWith('.xlsx') || file.originalname.endsWith('.xls')
      )
    })
    const chartContainer = ref(null)
    const chart = ref(null)
    const uploadRef = ref(null)
    const uploadHeaderRowStart = ref(0)
    const uploadHeaderRowEnd = ref(0)
    const uploadPassword = ref('')
    
    const loginForm = ref({
      username: '',
      password: ''
    })
    
    const registerForm = ref({
      username: '',
      email: '',
      password: ''
    })
    
    const compareForm = ref({
      file1Id: '',
      file2Id: '',
      sheet1: '',
      sheet2: '',
      primaryKeys: [],
      ignoreColumns: [],
      headerRowStart: 0,
      headerRowEnd: 0,
      nullValueStrategy: 'ignore',
      detectMoved: false,
      file1Password: '',
      file2Password: ''
    })
    
    const sheets1 = ref([])
    const sheets2 = ref([])
    const columns1 = ref([])
    const columns2 = ref([])
    const comparisonResult = ref(null)
    const previewDialogVisible = ref(false)
    const previewData = ref(null)
    const previewLoading = ref(false)
    const currentPreviewFileId = ref('')
    const previewNeedPassword = ref(false)
    const previewParams = ref({
      sheet: '',
      headerRowStart: 0,
      headerRowEnd: 0,
      limit: 30,
      password: ''
    })
    
    // 历史记录相关
    const historyForm = ref({
      search: '',
      dateRange: []
    })
    const historyList = ref([])
    const historyTotal = ref(0)
    const historyPage = ref(1)
    const historyLimit = ref(10)
    
    // 数据看板相关
    const dashboardStats = ref({
      totalFiles: 0,
      totalComparisons: 0,
      todayComparisons: 0,
      last7Days: [],
      differenceDistribution: {
        added: 0,
        deleted: 0,
        modified: 0
      }
    })
    const trendChart = ref(null)
    const distributionChart = ref(null)
    const timeChart = ref(null)
    
    // 检查登录状态
    const checkLoginStatus = () => {
      const storedToken = localStorage.getItem('token')
      const storedUser = localStorage.getItem('user')
      
      if (storedToken && storedUser) {
        token.value = storedToken
        user.value = JSON.parse(storedUser)
        isLoggedIn.value = true
        fetchDataList()
      }
    }
    
    // 登录
    const login = async () => {
      try {
        const response = await axios.post('http://localhost:5000/api/auth/login', loginForm.value)
        localStorage.setItem('token', response.data.token)
        localStorage.setItem('user', JSON.stringify(response.data.user))
        token.value = response.data.token
        user.value = response.data.user
        isLoggedIn.value = true
        fetchDataList()
      } catch (error) {
        console.error('登录错误:', error)
        alert('登录失败: ' + (error.response?.data?.msg || '服务器错误'))
      }
    }
    
    // 注册
    const register = async () => {
      try {
        const response = await axios.post('http://localhost:5000/api/auth/register', registerForm.value)
        localStorage.setItem('token', response.data.token)
        localStorage.setItem('user', JSON.stringify(response.data.user))
        token.value = response.data.token
        user.value = response.data.user
        isLoggedIn.value = true
        fetchDataList()
      } catch (error) {
        console.error('注册错误:', error)
        alert('注册失败: ' + (error.response?.data?.msg || '服务器错误'))
      }
    }
    
    // 退出登录
    const logout = () => {
      localStorage.removeItem('token')
      localStorage.removeItem('user')
      isLoggedIn.value = false
      user.value = {}
      token.value = ''
    }
    
    // 处理菜单选择
    const handleMenuSelect = (key) => {
      activeMenu.value = key
      if (key === 'visualization') {
        console.log('切换到数据可视化模块')
        fetchDataList()
      } else if (key === 'compare') {
        // 切换到对比页面时，确保重新获取数据列表
        fetchDataList()
      } else if (key === 'history') {
        fetchHistoryList()
      } else if (key === 'dashboard') {
        fetchDashboardStats()
      } else if (key === 'upload') {
        if (activeUploadTab.value === 'templateManagement') {
          fetchTemplateList()
        }
      }
    }
    
    // 监听 activeUploadTab 变化
    watch(activeUploadTab, (newVal) => {
      if (newVal === 'templateManagement') {
        fetchTemplateList()
      }
    })

    // 获取模板列表
    const fetchTemplateList = async () => {
      try {
        const response = await axios.get('http://localhost:5000/api/template/list', {
            headers: { 'x-auth-token': token.value }
        })
        templateList.value = response.data
      } catch (error) {
        console.error('获取模板列表失败:', error)
      }
    }

    // 模板上传处理
    const handleTemplateUploadSuccess = (response) => {
      // alert('模板上传成功')
      templateFileList.value = []
      templateUploadDialogVisible.value = false
      fetchTemplateList()
    }
    
    const handleTemplateUploadError = (error) => {
      console.error('模板上传失败:', error)
      alert('模板上传失败: ' + (error.response?.data?.msg || '服务器错误'))
    }

    const submitTemplateUpload = () => {
      if (templateUploadRef.value) {
        templateUploadRef.value.submit()
      }
    }

    // 删除模板
    const deleteTemplate = async (id) => {
      if (confirm('确定要删除这个模板吗？')) {
        try {
          await axios.delete(`http://localhost:5000/api/template/${id}`, {
            headers: { 'x-auth-token': token.value }
          })
          fetchTemplateList()
        } catch (error) {
          console.error('删除模板失败:', error)
          alert('删除模板失败')
        }
      }
    }

    // 下载模板
    const downloadTemplate = async (template) => {
      try {
        const response = await axios.get(`http://localhost:5000/api/template/download/${template._id}`, {
          headers: { 'x-auth-token': token.value },
          responseType: 'blob'
        })
        
        const url = window.URL.createObjectURL(new Blob([response.data]))
        const link = document.createElement('a')
        link.href = url
        link.setAttribute('download', template.originalname)
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
        window.URL.revokeObjectURL(url)
      } catch (error) {
        console.error('下载模板失败:', error)
        alert('下载模板失败')
      }
    }

    // 预览模板
    const previewTemplate = async (template) => {
      previewDialogVisible.value = true
      previewLoading.value = true
      previewData.value = null
      currentPreviewFileId.value = template._id
      previewNeedPassword.value = false
      isPreviewingTemplate.value = true
      
      previewParams.value = {
        sheet: '',
        headerRowStart: 0,
        headerRowEnd: 0,
        limit: 30,
        password: ''
      }
      
      await fetchPreviewData()
    }

    // 处理文件上传成功
    const handleUploadSuccess = (response) => {
      // alert('上传成功') // 移除成功提示
      fileList.value = []
      uploadPassword.value = '' // 清空密码
      fetchDataList()
    }
    
    // 处理文件上传失败
    const handleUploadError = (error) => {
      console.error('上传失败:', error)
      let errorMessage = '上传失败: '
      if (error.response && error.response.data && error.response.data.msg) {
        errorMessage += error.response.data.msg
      } else if (error.message) {
        errorMessage += error.message
      } else if (error.status) {
        errorMessage += '状态码: ' + error.status
      } else {
        errorMessage += '未知错误'
      }
      alert(errorMessage)
    }
    
    // 提交上传
    const submitUpload = () => {
      if (!token.value) {
        alert('请先登录')
        return
      }
      if (uploadRef.value) {
        uploadRef.value.submit()
      }
    }
    
    // 获取数据列表
    const fetchDataList = async () => {
      try {
        const response = await axios.get('http://localhost:5000/api/data/list', {
          headers: {
            'x-auth-token': token.value
          }
        })
        dataList.value = response.data
      } catch (error) {
        console.error('获取数据列表失败:', error)
      }
    }
    
    // 删除数据
    const deleteData = async (id) => {
      if (confirm('确定要删除这个文件吗？')) {
        try {
          console.log('删除文件ID:', id)
          await axios.delete(`http://localhost:5000/api/data/${id}`, {
            headers: {
              'x-auth-token': token.value
            }
          })
          fetchDataList()
          // alert('删除成功') // 移除成功提示
        } catch (error) {
          console.error('删除失败:', error)
          alert('删除失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      }
    }
    
    // 格式化文件大小
    const formatSize = (row, column) => {
      const size = row.size
      if (size < 1024) {
        return size + ' B'
      } else if (size < 1024 * 1024) {
        return (size / 1024).toFixed(2) + ' KB'
      } else {
        return (size / (1024 * 1024)).toFixed(2) + ' MB'
      }
    }
    
    // 格式化日期
    const formatDate = (row, column) => {
      const date = new Date(row.createdAt)
      return date.toLocaleString()
    }

    const isEncrypted = (row) => {
      return row.metadata && (row.metadata.encrypted || row.metadata.passwordProtected)
    }
    
    // 初始化图表
    const initChart = (retryCount = 0) => {
      // 确保容器存在
      if (!chartContainer.value) {
        console.warn('chartContainer is null, retrying...', retryCount)
        if (retryCount < 5) {
          setTimeout(() => {
            initChart(retryCount + 1)
          }, 200)
        } else {
          console.error('Failed to get chartContainer after 5 retries')
        }
        return
      }

      console.log('chartContainer found:', chartContainer.value)
      
      // 销毁旧实例
      if (chart.value) {
        chart.value.dispose()
        chart.value = null
      }
      
      try {
        chart.value = echarts.init(chartContainer.value)
        
        const option = {
          title: {
            text: '上传文件大小统计',
            left: 'center'
          },
          tooltip: {
            trigger: 'axis',
            axisPointer: {
              type: 'shadow'
            }
          },
          grid: {
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
          },
          xAxis: [
            {
              type: 'category',
              data: dataList.value.length > 0 
                ? dataList.value.slice(0, 15).map(item => item.originalname.length > 10 ? item.originalname.substring(0, 10) + '...' : item.originalname) 
                : ['示例A', '示例B', '示例C', '示例D', '示例E'],
              axisTick: {
                alignWithLabel: true
              },
              axisLabel: {
                interval: 0,
                rotate: 30
              }
            }
          ],
          yAxis: [
            {
              type: 'value',
              name: '大小 (KB)'
            }
          ],
          series: [
            {
              name: '文件大小',
              type: 'bar',
              barWidth: '60%',
              data: dataList.value.length > 0 
                ? dataList.value.slice(0, 15).map(item => (item.size / 1024).toFixed(2)) 
                : [10, 52, 200, 334, 390],
              itemStyle: {
                color: '#409EFF'
              }
            }
          ]
        }

        console.log('Setting chart option:', option)
        chart.value.setOption(option)
        
        // 监听窗口大小变化
        window.addEventListener('resize', () => {
          if (chart.value) {
            chart.value.resize()
          }
        })
      } catch (error) {
        console.error('图表初始化失败:', error)
      }
    }
    
    // 生成报表
    // const generateReport = () => {
    //   // alert('报表生成成功') // 移除成功提示
    // }
    
    // 导出报表
    // const exportReport = () => {
    //   // alert('报表导出成功') // 移除成功提示
    // }
    
    // 下载报表
    const downloadReport = async (taskId) => {
      try {
        const response = await axios.get(`http://localhost:5000/api/compare/report/${taskId}`, {
          headers: {
            'x-auth-token': token.value
          },
          responseType: 'blob'
        })
        
        // 创建下载链接
        const url = window.URL.createObjectURL(new Blob([response.data]))
        const link = document.createElement('a')
        link.href = url
        link.setAttribute('download', `report-${taskId}.xlsx`)
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
        window.URL.revokeObjectURL(url)
      } catch (error) {
        console.error('下载报表失败:', error)
        alert('下载报表失败')
      }
    }
    
    // 预览文件
    const previewFile = async (file) => {
      previewDialogVisible.value = true
      previewLoading.value = true
      previewData.value = null
      currentPreviewFileId.value = file._id
      previewNeedPassword.value = false
      isPreviewingTemplate.value = false
      
      // 重置参数
      previewParams.value = {
        sheet: '',
        headerRowStart: file.metadata?.headerRowStart || 0,
        headerRowEnd: file.metadata?.headerRowEnd !== undefined ? file.metadata.headerRowEnd : (file.metadata?.headerRowStart || 0),
        limit: 30,
        password: ''
      }
      
      await fetchPreviewData()
    }

    const fetchPreviewData = async () => {
      if (!currentPreviewFileId.value) return
      
      previewLoading.value = true
      try {
        const params = {
          headerRowStart: previewParams.value.headerRowStart,
          headerRowEnd: previewParams.value.headerRowEnd,
          limit: previewParams.value.limit
        }
        if (previewParams.value.sheet) {
          params.sheet = previewParams.value.sheet
        }
        if (previewParams.value.password) {
          params.password = previewParams.value.password
        }

        const baseUrl = isPreviewingTemplate.value ? 'http://localhost:5000/api/template/preview' : 'http://localhost:5000/api/data/preview'
        const response = await axios.get(`${baseUrl}/${currentPreviewFileId.value}`, {
          headers: {
            'x-auth-token': token.value
          },
          params
        })
        previewData.value = response.data
        previewNeedPassword.value = false // 成功获取数据，说明不需要密码或密码正确
        
        // 如果当前没有选中Sheet，则选中返回的Sheet
        if (!previewParams.value.sheet && response.data.sheet) {
          previewParams.value.sheet = response.data.sheet
        }
      } catch (error) {
        console.error('预览文件失败:', error)
        if (error.response?.data?.needPassword) {
          previewNeedPassword.value = true
          // 提示用户输入密码
          alert('该文件受密码保护，请输入密码')
        } else if (error.response?.status === 400 && error.response?.data?.msg?.includes('密码')) {
           previewNeedPassword.value = true
           alert(error.response.data.msg)
        } else {
           alert('预览文件失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      } finally {
        previewLoading.value = false
      }
    }

    const handlePreviewParamsChange = () => {
      fetchPreviewData()
    }

    // 下载文件
    const downloadFile = async (file) => {
      try {
        const response = await axios.get(`http://localhost:5000/api/data/download/${file._id}`, {
          headers: {
            'x-auth-token': token.value
          },
          responseType: 'blob'
        })
        
        // 创建下载链接
        const url = window.URL.createObjectURL(new Blob([response.data]))
        const link = document.createElement('a')
        link.href = url
        link.setAttribute('download', file.originalname)
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
        window.URL.revokeObjectURL(url)
      } catch (error) {
        console.error('下载文件失败:', error)
        alert('下载文件失败')
      }
    }
    
    // 处理第一个文件选择
    const handleFile1Change = async (fileId) => {
      if (fileId) {
        // 检查文件是否需要密码
        const file = dataList.value.find(f => f._id === fileId)
        compareForm.value.file1Password = '' // 重置密码
        
        // 自动填充表头范围
        if (file && file.metadata) {
          compareForm.value.headerRowStart = file.metadata.headerRowStart || 0
          compareForm.value.headerRowEnd = file.metadata.headerRowEnd !== undefined ? file.metadata.headerRowEnd : (file.metadata.headerRowStart || 0)
        }

        // 尝试获取sheet列表，如果需要密码会失败
        try {
          await fetchSheets1(fileId)
        } catch (error) {
          // fetchSheets1 内部会处理密码提示
        }
      } else {
        sheets1.value = []
        compareForm.value.sheet1 = ''
        columns1.value = []
        compareForm.value.primaryKeys = []
      }
    }
    
    const fetchSheets1 = async (fileId) => {
      try {
        const params = {}
        if (compareForm.value.file1Password) {
          params.password = compareForm.value.file1Password
        }
        
        const response = await axios.get(`http://localhost:5000/api/compare/sheets/${fileId}`, {
          headers: {
            'x-auth-token': token.value
          },
          params
        })
        sheets1.value = response.data.sheets
        compareForm.value.sheet1 = response.data.sheets[0] || ''
        // 获取列信息和推荐主键
        if (compareForm.value.sheet1) {
          await fetchColumns1(fileId, compareForm.value.sheet1)
        }
      } catch (error) {
        console.error('获取Sheet列表失败:', error)
        if (error.response?.data?.needPassword) {
           const password = prompt('文件1受密码保护，请输入密码:')
           if (password) {
             compareForm.value.file1Password = password
             await fetchSheets1(fileId)
           } else {
             // 用户取消输入密码，清除选择
             compareForm.value.file1Id = ''
             sheets1.value = []
           }
        } else {
           alert('获取文件1信息失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      }
    }

    // 处理第二个文件选择
    const handleFile2Change = async (fileId) => {
      if (fileId) {
        compareForm.value.file2Password = '' // 重置密码
        try {
          await fetchSheets2(fileId)
        } catch (error) {
          // fetchSheets2 内部会处理
        }
      } else {
        sheets2.value = []
        compareForm.value.sheet2 = ''
        columns2.value = []
      }
    }
    
    const fetchSheets2 = async (fileId) => {
      try {
        const params = {}
        if (compareForm.value.file2Password) {
          params.password = compareForm.value.file2Password
        }

        const response = await axios.get(`http://localhost:5000/api/compare/sheets/${fileId}`, {
          headers: {
            'x-auth-token': token.value
          },
          params
        })
        sheets2.value = response.data.sheets
        compareForm.value.sheet2 = response.data.sheets[0] || ''
      } catch (error) {
        console.error('获取Sheet列表失败:', error)
        if (error.response?.data?.needPassword) {
           const password = prompt('文件2受密码保护，请输入密码:')
           if (password) {
             compareForm.value.file2Password = password
             await fetchSheets2(fileId)
           } else {
             compareForm.value.file2Id = ''
             sheets2.value = []
           }
        } else {
           alert('获取文件2信息失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      }
    }
    
    // 处理第一个文件Sheet选择
    const handleSheet1Change = async (sheet) => {
      if (compareForm.value.file1Id && sheet) {
        await fetchColumns1(compareForm.value.file1Id, sheet)
      }
    }
    
    // 获取第一个文件的列信息和推荐主键
    const fetchColumns1 = async (fileId, sheet) => {
      try {
        const params = {
          headerRowStart: compareForm.value.headerRowStart,
          headerRowEnd: compareForm.value.headerRowEnd
        }
        if (compareForm.value.file1Password) {
          params.password = compareForm.value.file1Password
        }
        
        const response = await axios.get(`http://localhost:5000/api/compare/columns/${fileId}/${encodeURIComponent(sheet)}`, {
          headers: {
            'x-auth-token': token.value
          },
          params
        })
        columns1.value = response.data.columns
        // 默认选择所有列
        compareForm.value.primaryKeys = response.data.columns
      } catch (error) {
        console.error('获取列信息失败:', error)
        // 这里一般不需要再次处理密码，因为在获取Sheet时已经处理过了，除非Session过期等
      }
    }
    
    // 获取历史记录列表
    const fetchHistoryList = async () => {
      try {
        const params = {
          page: historyPage.value,
          limit: historyLimit.value,
          search: historyForm.value.search
        }
        
        if (historyForm.value.dateRange && historyForm.value.dateRange.length === 2) {
          params.startDate = historyForm.value.dateRange[0]
          params.endDate = historyForm.value.dateRange[1]
        }
        
        const response = await axios.get('http://localhost:5000/api/history/list', {
          headers: {
            'x-auth-token': token.value
          },
          params
        })
        
        historyList.value = response.data.data
        historyTotal.value = response.data.total
      } catch (error) {
        console.error('获取历史记录失败:', error)
      }
    }
    
    // 重置历史记录表单
    const resetHistoryForm = () => {
      historyForm.value = {
        search: '',
        dateRange: []
      }
      historyPage.value = 1
      fetchHistoryList()
    }
    
    // 查看历史记录
    const viewHistory = async (id) => {
      try {
        const response = await axios.get(`http://localhost:5000/api/history/${id}`, {
          headers: {
            'x-auth-token': token.value
          }
        })
        // 这里可以显示历史记录详情
        console.log('历史记录详情:', response.data)
      } catch (error) {
        console.error('获取历史记录详情失败:', error)
      }
    }
    
    // 处理分页大小变化
    const handleSizeChange = (size) => {
      historyLimit.value = size
      fetchHistoryList()
    }
    
    // 处理页码变化
    const handleCurrentChange = (current) => {
      historyPage.value = current
      fetchHistoryList()
    }
    
    // 获取数据看板数据
    const fetchDashboardStats = async () => {
      try {
        const response = await axios.get('http://localhost:5000/api/dashboard/stats', {
          headers: {
            'x-auth-token': token.value
          }
        })
        
        if (response.data.user) {
          // 管理员
          dashboardStats.value = response.data.user
        } else {
          // 普通用户
          dashboardStats.value = response.data
        }
        
        // 初始化图表
        initDashboardCharts()
      } catch (error) {
        console.error('获取仪表盘数据失败:', error)
      }
    }
    
    // 初始化数据看板图表
    const initDashboardCharts = () => {
      // 确保有数据，如果没有则使用默认数据防止空白
      const safeStats = dashboardStats.value || {
        last7Days: [],
        differenceDistribution: { added: 0, deleted: 0, modified: 0 }
      }
      
      // 趋势图
      if (trendChart.value) {
        if (echarts.getInstanceByDom(trendChart.value)) {
          echarts.getInstanceByDom(trendChart.value).dispose()
        }
        const chart = echarts.init(trendChart.value)
        const option = {
          tooltip: {
            trigger: 'axis'
          },
          grid: {
            top: '10%',
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
          },
          xAxis: {
            type: 'category',
            data: safeStats.last7Days.length > 0 ? safeStats.last7Days.map(item => item.date) : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
          },
          yAxis: {
            type: 'value'
          },
          series: [{
            data: safeStats.last7Days.length > 0 ? safeStats.last7Days.map(item => item.count) : [0, 0, 0, 0, 0, 0, 0],
            type: 'line',
            smooth: true,
            areaStyle: {
              opacity: 0.3
            },
            itemStyle: {
              color: '#409EFF'
            }
          }]
        }
        chart.setOption(option)
      }
      
      // 差异类型分布图
      if (distributionChart.value) {
        if (echarts.getInstanceByDom(distributionChart.value)) {
          echarts.getInstanceByDom(distributionChart.value).dispose()
        }
        const chart = echarts.init(distributionChart.value)
        const hasData = safeStats.differenceDistribution.added > 0 || 
                        safeStats.differenceDistribution.deleted > 0 || 
                        safeStats.differenceDistribution.modified > 0
                        
        const option = {
          tooltip: {
            trigger: 'item'
          },
          legend: {
            top: '5%',
            left: 'center'
          },
          series: [{
            name: '差异类型',
            type: 'pie',
            radius: ['40%', '70%'],
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
                fontSize: '18',
                fontWeight: 'bold'
              }
            },
            labelLine: {
              show: false
            },
            data: hasData ? [
              { value: safeStats.differenceDistribution.added, name: '新增', itemStyle: { color: '#67C23A' } },
              { value: safeStats.differenceDistribution.deleted, name: '删除', itemStyle: { color: '#F56C6C' } },
              { value: safeStats.differenceDistribution.modified, name: '修改', itemStyle: { color: '#E6A23C' } }
            ] : [
              { value: 1, name: '无数据', itemStyle: { color: '#909399' } }
            ]
          }]
        }
        chart.setOption(option)
      }
      
      // 任务时间趋势图
      if (timeChart.value) {
        if (echarts.getInstanceByDom(timeChart.value)) {
          echarts.getInstanceByDom(timeChart.value).dispose()
        }
        const chart = echarts.init(timeChart.value)
        const option = {
          tooltip: {
            trigger: 'axis',
            axisPointer: {
              type: 'shadow'
            }
          },
          grid: {
            top: '10%',
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
          },
          xAxis: {
            type: 'category',
            data: safeStats.last7Days.length > 0 ? safeStats.last7Days.map(item => item.date) : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
          },
          yAxis: {
            type: 'value'
          },
          series: [{
            data: safeStats.last7Days.length > 0 ? safeStats.last7Days.map(item => item.count) : [0, 0, 0, 0, 0, 0, 0],
            type: 'bar',
            barWidth: '40%',
            itemStyle: {
              color: '#409EFF',
              borderRadius: [5, 5, 0, 0]
            }
          }]
        }
        chart.setOption(option)
      }
      
      // 监听窗口大小变化
      window.addEventListener('resize', () => {
        [trendChart.value, distributionChart.value, timeChart.value].forEach(container => {
          if (container) {
            const instance = echarts.getInstanceByDom(container)
            if (instance) instance.resize()
          }
        })
      })
    }
    
    // 监听 analysisResult 变化，如果清空了，销毁图表
    watch(comparisonResult, (newVal) => {
      // 可以在这里添加对比结果变化时的逻辑
    })
    
    // 对比文件
    const compareFiles = async () => {
      try {
        const response = await axios.post('http://localhost:5000/api/compare/compare', compareForm.value, {
          headers: {
            'x-auth-token': token.value
          }
        })
        comparisonResult.value = response.data
      } catch (error) {
        console.error('文件对比失败:', error)
        if (error.response?.data?.needFile1Password) {
           const password = prompt(`文件1 "${error.response.data.msg}" 受密码保护，请输入密码:`)
           if (password) {
             compareForm.value.file1Password = password
             compareFiles() // 重试
           }
        } else if (error.response?.data?.needFile2Password) {
           const password = prompt(`文件2 "${error.response.data.msg}" 受密码保护，请输入密码:`)
           if (password) {
             compareForm.value.file2Password = password
             compareFiles() // 重试
           }
        } else {
           alert('文件对比失败: ' + (error.response?.data?.msg || '服务器错误'))
        }
      }
    }
    
    // 格式化行数据
    const formatRowData = (row, column) => {
      if (row.rowData) {
        return JSON.stringify(row.rowData)
      }
      return ''
    }
    
    // 格式化差异类型
    const formatDifferenceType = (row, column) => {
      const typeMap = {
        'added': '新增',
        'deleted': '删除',
        'modified': '修改',
        'moved': '移动'
      }
      return typeMap[row.type] || row.type
    }
    
    onMounted(() => {
      checkLoginStatus()
    })
    
    return {
      isLoggedIn,
      user,
      token,
      activeMenu,
      activeTab,
      fileList,
      dataList,
      excelFiles,
      loginLoading,
      registerLoading,
      chartContainer,
      uploadRef,
      uploadHeaderRowStart,
      uploadHeaderRowEnd,
      uploadPassword,
      loginForm,
      registerForm,
      compareForm,
      sheets1,
      sheets2,
      columns1,
      columns2,
      comparisonResult,
      // 历史记录相关
      historyForm,
      historyList,
      historyTotal,
      historyPage,
      historyLimit,
      // 数据看板相关
      dashboardStats,
      trendChart,
      distributionChart,
      timeChart,
      login,
      register,
      logout,
      handleMenuSelect,
      handleUploadSuccess,
      handleUploadError,
      submitUpload,
      deleteData,
      formatSize,
      formatDate,
      handleFile1Change,
      handleFile2Change,
      handleSheet1Change,
      compareFiles,
      formatRowData,
      formatDifferenceType,
      downloadReport,
      previewFile,
      previewDialogVisible,
      previewData,
      previewLoading,
      previewParams,
      previewNeedPassword,
      handlePreviewParamsChange,
      downloadFile,
      isEncrypted,
      templateList,
      templateUploadDialogVisible,
      templateFileList,
      templateUploadRef,
      activeUploadTab,
      fetchTemplateList,
      handleTemplateUploadSuccess,
      handleTemplateUploadError,
      submitTemplateUpload,
      deleteTemplate,
      downloadTemplate,
      previewTemplate,
      
      // 历史记录方法
      fetchHistoryList,
      resetHistoryForm,
      viewHistory,
      handleSizeChange,
      handleCurrentChange,
      // 数据看板方法
      fetchDashboardStats
    }
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  color: #2c3e50;
  min-height: 100vh;
}

.el-header {
  background-color: #409EFF;
  color: white;
  height: 60px;
  line-height: 60px;
}

.header-content {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 0 20px;
}

.header-content h1 {
  margin: 0;
  font-size: 20px;
}

.user-info {
  display: flex;
  align-items: center;
  gap: 10px;
}

.el-aside {
  background-color: #f0f2f5;
  height: calc(100vh - 60px);
}

.el-main {
  padding: 20px;
  background-color: #f9f9f9;
}

.login-container {
  height: 100vh;
  display: flex;
  justify-content: center;
  align-items: center;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
}

.login-box {
  width: 400px;
  background: white;
  padding: 40px;
  border-radius: 12px;
  box-shadow: 0 15px 35px rgba(0, 0, 0, 0.2);
}

.login-header {
  text-align: center;
  margin-bottom: 30px;
}

.login-logo {
  width: 60px;
  margin-bottom: 10px;
}

.login-header h2 {
  margin: 10px 0 5px;
  color: #333;
  font-size: 28px;
  font-weight: 600;
}

.login-header p {
  color: #666;
  margin: 0;
  font-size: 14px;
}

.login-btn {
  width: 100%;
  height: 45px;
  font-size: 16px;
  margin-top: 10px;
}

.el-tabs__item {
  font-size: 16px;
  height: 45px;
}

.el-form-item__label {
  font-weight: 500;
}

.el-upload__tip {
  margin-top: 10px;
}

.preview-content {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.mb-4 {
  margin-bottom: 16px;
}

.preview-toolbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 15px;
  background-color: #f8f9fa;
  padding: 12px;
  border-radius: 6px;
  border: 1px solid #ebeef5;
}

.toolbar-left, .toolbar-right {
  display: flex;
  align-items: center;
  gap: 15px;
  flex-wrap: wrap;
}

.toolbar-item {
  display: flex;
  align-items: center;
  gap: 8px;
}

.toolbar-item .label {
  font-size: 14px;
  color: #606266;
  white-space: nowrap;
}

.error-text {
  color: #f56c6c !important;
  font-weight: bold;
}

.empty-preview {
  height: 400px;
  display: flex;
  justify-content: center;
  align-items: center;
}
</style>
