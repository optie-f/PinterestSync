import { RecordPinsData } from './pinterestService';
import { SaveToDrive } from './driveService';

declare var global: any;

global.recordPins = (): void => {
  RecordPinsData.main();
};
global.saveToDrive = (): void => {
  SaveToDrive.main();
};
