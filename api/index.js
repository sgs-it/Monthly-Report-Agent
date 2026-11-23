import worker from '../dist/server/index.js';
import * as storage from '@vercel/blob';
import {handleUpload} from '@vercel/blob/client';
import {createHandler} from '../server/vercel-handler.js';
const handler=createHandler({worker,storage,handleUpload,env:process.env});
export default {fetch:handler};
