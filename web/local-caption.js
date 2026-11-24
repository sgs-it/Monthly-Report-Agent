import {PREDEFINED_CAPTIONS} from '../src/captions.js';

// Keep capitalization variants in the picker but avoid competing identical labels during inference.
const labels = [...new Map(PREDEFINED_CAPTIONS.filter(Boolean).map(label => [label.toLowerCase(), label])).values()];
let classifierPromise;

async function loadClassifier() {
  const {pipeline, env} = await import('https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.2/dist/transformers.min.js');
  env.allowLocalModels = false;
  env.useBrowserCache = true;
  return pipeline('zero-shot-image-classification', 'Xenova/clip-vit-base-patch32', {quantized:true});
}

export async function suggestCaption(photoId) {
  classifierPromise ??= loadClassifier().catch(error => {
    classifierPromise = undefined;
    throw error;
  });
  const classify = await classifierPromise;
  const response = await fetch('/api/images/' + encodeURIComponent(photoId));
  if (!response.ok) throw Error('Could not read the saved photo.');
  const imageURL = URL.createObjectURL(await response.blob());
  try {
    const [best] = await classify(imageURL, labels);
    return labels.includes(best?.label) ? best.label : '';
  } finally {
    URL.revokeObjectURL(imageURL);
  }
}
