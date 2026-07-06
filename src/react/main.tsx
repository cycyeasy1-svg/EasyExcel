import './polyfills/buffer';
import ReactDOM from 'react-dom/client';
import './util/vscode.js';
import './main.css';
import { ConfigProvider } from 'antd';
import { antThemeConfig } from './antThemeConfig.ts';
import Excel from './view/excel/Excel.tsx';

document.getElementById('_defaultStyles')?.parentNode?.removeChild(document.getElementById('_defaultStyles'));

export default function App() {
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
