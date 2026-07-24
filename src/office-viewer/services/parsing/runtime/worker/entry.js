// 独立 JS 入口确保源码开发和 Father bundless 产物使用同一个稳定 URL。
import { runOfficeParserWorker } from './runOfficeParserWorker';

runOfficeParserWorker(self);
