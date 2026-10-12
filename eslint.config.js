// ESLint 扁平配置。三套运行环境差异很大，所以按目录分开声明，别用一套兜住：
//   server/taskpane.js + server/actions/*  → 浏览器脚本（窗格），有 Office 全局
//   server/server.js                      → Node CommonJS（桥接服务）
//   lib/** + test/** + 本文件              → Node ESM
import js from '@eslint/js';
import globals from 'globals';

// Office.js 注入的宿主对象模型：只在加载项页面里存在，运行时由 Office 提供
const OFFICE_GLOBALS = {
  Office: 'readonly',
  Word: 'readonly',
  Excel: 'readonly',
  PowerPoint: 'readonly',
  // 由 /pako.min.js 以 <script> 引入，解压 OOXML 用
  pako: 'readonly',
};

const COMMON_RULES = {
  // 分片里大量「尽力而为」的 catch 只写注释，空块是刻意的
  'no-empty': ['error', { allowEmptyCatch: true }],
  'no-unused-vars': ['error', { args: 'after-used', argsIgnorePattern: '^_', caughtErrors: 'none' }],
  // 宿主回调经常需要 void 返回值
  'no-void': 'off',
};

export default [
  {
    ignores: ['node_modules/**', 'server/pako.min.js', '.tmp/**', 'dist/**'],
  },
  js.configs.recommended,

  {
    // 窗格侧：浏览器脚本，非模块（服务端拼接后作为单个 <script> 下发）
    files: ['server/taskpane.js', 'server/actions/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: { ...globals.browser, ...OFFICE_GLOBALS },
    },
    rules: COMMON_RULES,
  },

  {
    // 设置面板客户端：跑在 harness 的 web UI 里，是浏览器环境
    files: ['lib/client.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: { ...globals.browser, ...OFFICE_GLOBALS },
    },
    rules: COMMON_RULES,
  },

  {
    // 桥接服务：Node CommonJS
    files: ['server/server.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'commonjs',
      globals: { ...globals.node },
    },
    rules: COMMON_RULES,
  },

  {
    // 宿主侧入口 + 测试 + 本配置
    files: ['lib/index.js', 'lib/sidecar-host.mjs', 'test/**/*.mjs', 'eslint.config.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: { ...globals.node },
    },
    rules: COMMON_RULES,
  },
];
