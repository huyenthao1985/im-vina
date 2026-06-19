const { createClient } = require('@supabase/supabase-js');

const supabaseUrl = 'https://pquxjrfyafsaybuzovqy.supabase.co';
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo';

const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function test() {
  console.log('Testing connection...');
  const { data, error } = await supabase.from('sales_data').select('*').limit(1);
  if (error) {
    console.error('Error:', error);
  } else {
    console.log('Success, data:', data);
  }
}

test();
