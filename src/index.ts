import { RecordPinsData } from './pinterestService';
import { SaveToDrive } from './driveService';

declare var global: any;

global.recordPins = RecordPinsData.main();
global.saveToDrive = SaveToDrive.main();
