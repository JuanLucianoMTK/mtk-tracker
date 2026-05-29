import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = "https://sjcoyrfixyojvximfeur.supabase.co"
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNqY295cmZpeHlvanZ4aW1mZXVyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk1NzYyMTMsImV4cCI6MjA5NTE1MjIxM30.7IuqldX23vIUx8hh-NOKCad4TBsRxuslADq6Vez5DUs"

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)
