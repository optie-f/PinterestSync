import { PinterestSync } from './callapi';

declare var global: any;

global.main = (): void => {
  PinterestSync.main();
};
