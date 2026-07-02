const { spawn } = require('child_process');

function addEnv(key, value) {
  return new Promise((resolve, reject) => {
    const child = spawn('npx', ['vercel', 'env', 'add', key, 'production', '--scope', 'xuan-thao-s-projects'], { shell: true });
    
    child.stdout.on('data', (data) => {
      const output = data.toString();
      console.log(output);
      if (output.includes('What’s the value of')) {
        child.stdin.write(value + '\n');
      }
    });

    child.stderr.on('data', (data) => {
      console.error(data.toString());
    });

    child.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`Failed with code ${code}`));
    });
  });
}

async function main() {
  try {
    console.log('Adding URL...');
    await addEnv('VITE_SUPABASE_URL', 'https://pquxjrfyafsaybuzovqy.supabase.co');
    console.log('Adding KEY...');
    await addEnv('VITE_SUPABASE_ANON_KEY', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo');
    console.log('Done!');
  } catch (err) {
    console.error(err);
  }
}

main();
