<template>
  <div class="app-container">
    <h2 class="title">Word 文件批量极速合并工具</h2>

    <div class="control-panel">
      <button class="primary-btn" @click="chooseFolder" :disabled="isMerging || isScanning">
        {{ isScanning ? '扫描中...' : '选择文件夹' }}
      </button>
      <span class="path-text" :title="folderPath">{{ folderPath || '等待选择文件夹...' }}</span>
    </div>

    <div v-if="isScanning" class="loading-overlay">
      <div class="spinner"></div>
      <p>正在深度扫描本地 Word 文件，请稍候...</p>
    </div>

    <div class="list-container" v-if="files.length > 0 && !isScanning">
      <div class="list-header">
        <span>读取到的文件 (按住同行拖拽排序，点击按钮可平滑跟随)</span>
      </div>

      <div class="file-list" ref="fileList" @dragover="autoScroll">
        <div v-for="(file, index) in files"
             :key="file.path"
             :id="'row-' + index"
             class="file-item"
             :class="{
               'is-dragging': dragIndex === index,
               'is-moved': movedIndex === index
             }"
             :draggable="!file.isEditing && !isMerging"
             @dragstart="dragStart(index, $event)"
             @dragenter="dragEnter(index, $event)"
             @dragover.prevent
             @dragend="dragEnd($event)">

          <div class="file-info">
            <span class="drag-handle" title="按住此处或整行即可拖拽排序">⋮⋮</span>

            <input type="checkbox" :id="'file-'+index" v-model="file.selected" :disabled="isMerging">
            <span class="file-index" :class="{ 'unselected-text': !file.selected }">{{ index + 1 }}. </span>

            <input v-if="file.isEditing"
                   type="text"
                   v-model="file.displayName"
                   @blur="finishEdit(file)"
                   @keyup.enter="finishEdit(file)"
                   class="edit-input"
                   ref="editInputs">

            <label v-else
                   :for="'file-'+index"
                   :class="{ 'unselected-text': !file.selected }"
                   @dblclick="startEdit(file, index)"
                   title="双击可修改合并后的目录名">
              {{ file.displayName }}
            </label>
          </div>

          <div class="file-actions">
            <button class="icon-btn up-btn" @click.stop="moveUp(index)" :disabled="index === 0 || isMerging" title="上移">▲</button>
            <button class="icon-btn down-btn" @click.stop="moveDown(index)" :disabled="index === files.length - 1 || isMerging" title="下移">▼</button>
          </div>
        </div>
      </div>
    </div>

    <div v-if="isMerging" class="progress-container">
      <div class="progress-text">{{ progressMsg }} ({{ progressPercent }}%)</div>
      <div class="progress-bar-bg">
        <div class="progress-bar-fill" :style="{ width: progressPercent + '%' }"></div>
      </div>
    </div>

    <div class="footer-panel" v-if="files.length > 0 && !isScanning">
      <div class="left-actions">
        <button class="secondary-btn" @click="selectAll(true)" :disabled="isMerging">全选</button>
        <button class="secondary-btn" @click="selectAll(false)" :disabled="isMerging">全不选</button>
      </div>
      <button class="success-btn" @click="executeMerge" :disabled="isMerging">
        {{ isMerging ? '合并中，请稍候...' : '合并选定的文件' }}
      </button>
    </div>
  </div>
</template>

<script>
import { nextTick } from 'vue';

export default {
  data() {
    return {
      folderPath: '',
      files: [],
      isMerging: false,
      isScanning: false,
      progressPercent: 0,
      progressMsg: '',
      dragIndex: null,
      // 新增：记录刚才被点击移动的行号，用于展示高亮
      movedIndex: null
    }
  },
  mounted() {
    window.eel.expose(this.updateProgress, 'update_progress');
    window.eel.expose(this.setScanningStatus, 'set_scanning_status');
  },
  methods: {
    getFileName(fullPath) {
      const nameWithExt = fullPath.split('\\').pop().split('/').pop();
      return nameWithExt.replace(/\.docx$/i, '');
    },
    updateProgress(percent, msg) {
      this.progressPercent = Math.round(percent);
      this.progressMsg = msg;
    },
    setScanningStatus(status) {
      this.isScanning = status;
    },

    async chooseFolder() {
      const result = await window.eel.py_choose_and_scan()();
      this.isScanning = false;
      if (result && result.folder_path) {
        this.folderPath = result.folder_path;
        this.files = result.files.map(f => ({
          path: f,
          selected: true,
          displayName: this.getFileName(f),
          isEditing: false
        }));
      }
    },

    startEdit(file, index) {
      if (this.isMerging) return;
      file.isEditing = true;
      nextTick(() => {
        const inputs = this.$refs.editInputs;
        if (inputs && inputs[0]) inputs[0].focus();
      });
    },

    finishEdit(file) {
      file.isEditing = false;
      if (!file.displayName.trim()) {
        file.displayName = this.getFileName(file.path);
      }
    },

    // ================= 原生拖拽排序逻辑 =================
    dragStart(index, event) {
      this.dragIndex = index;
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', index);

      setTimeout(() => {
        if(event.target && event.target.classList) {
          event.target.classList.add('dragging-ghost');
        }
      }, 0);
    },

    dragEnter(index, event) {
      event.preventDefault();
      if (this.dragIndex === null || this.dragIndex === index) return;

      const draggedItem = this.files[this.dragIndex];
      this.files.splice(this.dragIndex, 1);
      this.files.splice(index, 0, draggedItem);
      this.dragIndex = index;
    },

    autoScroll(event) {
      event.preventDefault();
      const container = this.$refs.fileList;
      if (!container || this.dragIndex === null) return;

      const rect = container.getBoundingClientRect();
      const threshold = 50;
      const scrollSpeed = 15;

      if (event.clientY - rect.top < threshold) {
        container.scrollTop -= scrollSpeed;
      } else if (rect.bottom - event.clientY < threshold) {
        container.scrollTop += scrollSpeed;
      }
    },

    dragEnd(event) {
      this.dragIndex = null;
      if(event.target && event.target.classList) {
        event.target.classList.remove('dragging-ghost');
      }
    },

    selectAll(status) { this.files.forEach(f => f.selected = status); },

    // ================= 核心新增：视野追踪与高亮机制 =================
    trackFocus(newIndex) {
      // 1. 设置视觉高亮
      this.movedIndex = newIndex;
      // 0.8秒后自动取消高亮
      setTimeout(() => {
        if (this.movedIndex === newIndex) {
          this.movedIndex = null;
        }
      }, 800);

      // 2. 核心平滑滚动逻辑 (等 Vue 数据渲染到 DOM 后执行)
      nextTick(() => {
        const el = document.getElementById('row-' + newIndex);
        if (el) {
          // behavior: 'smooth' 开启平滑滚动动画
          // block: 'nearest' 表示如果元素跑出去了，就把它滚动到边缘刚好能看见的位置，不会引起剧烈的跳动
          el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
      });
    },

    moveUp(index) {
      if (index > 0) {
        const temp = this.files[index - 1];
        this.files[index - 1] = this.files[index];
        this.files[index] = temp;

        // 调用追踪方法，并传入它最新的索引位置
        this.trackFocus(index - 1);
      }
    },

    moveDown(index) {
      if (index < this.files.length - 1) {
        const temp = this.files[index + 1];
        this.files[index + 1] = this.files[index];
        this.files[index] = temp;

        // 调用追踪方法，并传入它最新的索引位置
        this.trackFocus(index + 1);
      }
    },
    // ==============================================================

    async executeMerge() {
      const selectedItems = this.files.filter(f => f.selected).map(f => ({
        path: f.path,
        displayName: f.displayName
      }));

      if (selectedItems.length === 0) {
        alert("⚠️ 请至少保留一个需要合并的文件！");
        return;
      }

      const savePath = await window.eel.py_choose_save_path()();
      if (!savePath) return;

      this.isMerging = true;
      this.progressPercent = 0;
      this.progressMsg = '正在初始化合并环境...';

      const res = await window.eel.py_merge_files(selectedItems, savePath)();
      this.isMerging = false;

      if (res.status === 'success') {
        alert(res.msg);
      } else {
        alert("❌ 合并出错: " + res.msg);
      }
    }
  }
}
</script>

<style scoped>
/* 样式基础 */
.app-container { max-width: 600px; margin: 20px auto; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
.title { color: #2c3e50; text-align: center; }
.control-panel, .footer-panel { display: flex; justify-content: space-between; align-items: center; margin: 15px 0; }
.path-text { color: #666; font-size: 0.9em; flex-grow: 1; margin-left: 10px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.loading-overlay { display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 40px; background: #fdfdfd; border: 1px dashed #409eff; border-radius: 8px; margin: 15px 0; color: #409eff; font-weight: bold; }
.spinner { border: 4px solid rgba(64, 158, 255, 0.2); border-top: 4px solid #409eff; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin-bottom: 15px; }
@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

/* 列表区 */
.list-container { border: 1px solid #ddd; border-radius: 8px; overflow: hidden; margin-bottom: 15px; background: #fff;}
.list-header { background: #f5f7fa; padding: 10px; font-weight: bold; font-size: 0.9em; border-bottom: 1px solid #ddd; }
.file-list { max-height: 350px; overflow-y: auto; padding: 5px 10px; user-select: none; scroll-behavior: smooth; /* 增加全局原生平滑滚动 */}

/* 列表行样式 */
.file-item {
  padding: 8px 5px;
  border-bottom: 1px dashed #eee;
  display: flex;
  justify-content: space-between;
  align-items: center;
  background-color: #fff;
  transition: transform 0.1s ease, background-color 0.4s ease; /* 增加背景颜色过渡动画 */
  cursor: grab;
}
.file-item:hover { background-color: #f0f7ff; }
.file-item:last-child { border-bottom: none; }
.file-item:active { cursor: grabbing; }

/* ================= 拖拽与高亮样式 ================= */
.dragging-ghost { opacity: 0.4; background-color: #e6f7ff; border: 1px dashed #1890ff; }
/* 按钮点击移动后的高亮呼吸效果 */
.is-moved { background-color: #e6f7ff !important; border-radius: 4px; box-shadow: inset 0 0 5px rgba(24,144,255,0.2);}
.drag-handle { color: #ccc; font-weight: bold; margin-right: 8px; cursor: grab; font-size: 1.1em; letter-spacing: -2px;}
.drag-handle:hover { color: #409eff; }
/* ================================================== */

.file-info { display: flex; align-items: center; gap: 8px; flex-grow: 1; overflow: hidden; }
.file-index { color: #666; font-weight: 500;}
.edit-input { flex-grow: 1; padding: 2px 5px; border: 1px solid #409eff; border-radius: 3px; font-family: inherit; font-size: 0.95em; outline: none; box-shadow: 0 0 3px rgba(64,158,255,0.3); user-select: text; }
.file-info label { cursor: pointer; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding: 2px 5px;}
.unselected-text { text-decoration: line-through; color: #aaa !important; }

/* 按钮与进度条 */
.file-actions { display: flex; gap: 5px; }
.icon-btn { cursor: pointer; padding: 4px 8px; border: 1px solid #ddd; border-radius: 4px; background: white; color: #666; font-size: 0.8em; transition: 0.2s; }
.icon-btn:hover:not(:disabled) { background: #e0f2f1; border-color: #26a69a; color: #00897b; }
.icon-btn:disabled { opacity: 0.3; cursor: not-allowed; }
button { cursor: pointer; padding: 8px 16px; border: none; border-radius: 4px; font-weight: bold; transition: 0.2s;}
button:disabled { cursor: not-allowed; opacity: 0.6; }
.primary-btn { background: #409eff; color: white; }
.primary-btn:hover:not(:disabled) { background: #66b1ff; }
.secondary-btn { background: #f4f4f5; color: #909399; border: 1px solid #d3d4d6; }
.secondary-btn:hover:not(:disabled) { background: #e9e9eb; }
.success-btn { background: #67c23a; color: white; font-size: 1.05em; }
.success-btn:hover:not(:disabled) { background: #85ce61; }
.progress-container { margin: 15px 0; padding: 15px; background: #fdfdfd; border-radius: 6px; border: 1px solid #eee; }
.progress-text { font-size: 0.9em; color: #409eff; margin-bottom: 8px; font-weight: bold;}
.progress-bar-bg { width: 100%; height: 12px; background-color: #ebeef5; border-radius: 6px; overflow: hidden; }
.progress-bar-fill { height: 100%; background-color: #409eff; transition: width 0.2s ease; }
</style>