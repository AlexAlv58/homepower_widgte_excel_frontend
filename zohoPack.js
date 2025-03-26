import AdmZip from 'adm-zip';
import { promises as fs } from 'fs';
import path from 'path';

const config = {
  buildDir: './dist',
  outputDir: './build',
  outputFile: 'app.zip',
  overwrite: true
};

async function createZipArchive(options = {}) {
  const settings = { ...config, ...options };
  const outputPath = path.join(settings.outputDir, settings.outputFile);

  try {
    try {
      await fs.access(settings.outputDir);
    } catch {
      console.log(`Creating output directory: ${settings.outputDir}`);
      await fs.mkdir(settings.outputDir, { recursive: true });
    }

    try {
      await fs.access(outputPath);
      if (!settings.overwrite) {
        throw new Error(`Output file ${outputPath} already exists and overwrite is set to false`);
      }
      console.log(`Output file ${outputPath} already exists. Overwriting...`);
    } catch (err) {
    }

    const zip = new AdmZip();
    zip.addLocalFolder(settings.buildDir);

    console.log(`Adding folder ${settings.buildDir} to archive...`);
    const zipEntries = zip.getEntries();
    console.log(`Archive contains ${zipEntries.length} entries`);

    zip.writeZip(outputPath);

    const stats = await fs.stat(outputPath);
    const fileSizeInMB = (stats.size / (1024 * 1024)).toFixed(2);

    console.log(`✅ Created ${outputPath} successfully (${fileSizeInMB} MB)`);
  } catch (error) {
    console.error(`❌ Error creating zip archive: ${error.message}`);
    process.exit(1);
  }
}


createZipArchive();


export { createZipArchive };
