import './polyfills/buffer';
import { lazy, Suspense } from 'react';
import ReactDOM from 'react-dom/client';
import './util/vscode.js';
import './main.css';
import { ConfigProvider } from 'antd';
import { antThemeConfig } from './antThemeConfig.ts';
import { getConfigs } from './util/vscodeConfig.ts';
import Excel from './view/excel/Excel.tsx';

document.getElementById('_defaultStyles')?.parentNode?.removeChild(document.getElementById('_defaultStyles'));

// Phase 0 spike entry (?univer-spike): lazy so the Univer chunk stays out of
// the default load path.
const ExcelDiff = lazy(() => import('./view/excel/ExcelDiff.tsx'));
const UniverSpike = lazy(() => import('./view/excel/univer/UniverSpike.tsx'));
const useUniverSpike = new URLSearchParams(window.location.search).has('univer-spike');

export default function App() {
  if (getConfigs()?.route === 'excel-diff') {
    return (
      <Suspense fallback={<div>loading excel diff...</div>}>
        <ExcelDiff />
      </Suspense>
    );
  }
  if (useUniverSpike) {
    return (
      <Suspense fallback={<div>loading univer spike...</div>}>
        <UniverSpike />
      </Suspense>
    );
  }
  return <Excel />;
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <ConfigProvider
    componentSize='small'
    theme={antThemeConfig}
  >
    <App />
  </ConfigProvider>
);
