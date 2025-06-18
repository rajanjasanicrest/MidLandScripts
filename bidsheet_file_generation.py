#!/usr/bin/env python3
"""
Master script to run multiple Python files in sequence
"""

import subprocess
import sys
import os
from datetime import datetime

def print_separator(message):
    """Print a formatted separator with message"""
    print("\n" + "="*60)
    print(f" {message}")
    print("="*60 + "\n")

def run_python_file(file_path):
    """
    Run a Python file and handle its execution
    """
    if not os.path.exists(file_path):
        print(f"❌ ERROR: File '{file_path}' not found!")
        return False
    
    print(f"🚀 Starting execution of: {file_path}")
    start_time = datetime.now()
    
    try:
        # Run the Python file using subprocess
        result = subprocess.run([sys.executable, file_path], 
                              capture_output=True, 
                              text=True, 
                              check=True)
        
        # Print output if any
        if result.stdout:
            print("📤 Output:")
            print(result.stdout)
        
        end_time = datetime.now()
        execution_time = (end_time - start_time).total_seconds()
        
        print(f"✅ Successfully completed: {file_path}")
        print(f"⏱️  Execution time: {execution_time:.2f} seconds")
        return True
        
    except subprocess.CalledProcessError as e:
        end_time = datetime.now()
        execution_time = (end_time - start_time).total_seconds()
        
        print(f"❌ ERROR in {file_path}:")
        print(f"Exit code: {e.returncode}")
        if e.stdout:
            print("Output:", e.stdout)
        if e.stderr:
            print("Error:", e.stderr)
        print(f"⏱️  Execution time: {execution_time:.2f} seconds")
        return False
    
    except Exception as e:
        end_time = datetime.now()
        execution_time = (end_time - start_time).total_seconds()
        
        print(f"❌ Unexpected error running {file_path}: {str(e)}")
        print(f"⏱️  Execution time: {execution_time:.2f} seconds")
        return False

def main():
    """
    Main function to run all Python files
    """
    # List of Python files to run (modify these names according to your files)
    python_files = [
        "bidsheet_brass_consolidate.py",          # Replace with your first Python file name
        "bidsheet_other_metal_bids_consolidate.py",          # Replace with your second Python file name
        "bidsheet_steel_consolidation.py", 
        "discount_rebate_consolidation.py", 
        "new_product_intro_consolidation.py", 
        "supply_chain_consolidation.py"           # Replace with your third Python file name
    ]
    
    print_separator("MASTER PYTHON RUNNER STARTED")
    print(f"📅 Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📝 Files to execute: {len(python_files)}")
    
    # Track execution results
    results = {}
    total_start_time = datetime.now()
    
    # Run each Python file
    for i, file_path in enumerate(python_files, 1):
        print_separator(f"EXECUTING FILE {i}/{len(python_files)}: {file_path}")
        success = run_python_file(file_path)
        results[file_path] = success
        
        if not success:
            print(f"\n⚠️  Warning: {file_path} failed to execute properly")
            
            # Ask user if they want to continue
            user_input = input("\nDo you want to continue with the next file? (y/n): ").lower().strip()
            if user_input not in ['y', 'yes']:
                print("🛑 Execution stopped by user")
                break
    
    # Summary
    total_end_time = datetime.now()
    total_execution_time = (total_end_time - total_start_time).total_seconds()
    
    print_separator("EXECUTION SUMMARY")
    print(f"📅 Completed at: {total_end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️  Total execution time: {total_execution_time:.2f} seconds")
    print(f"📊 Results:")
    
    successful_count = 0
    for file_path, success in results.items():
        status = "✅ SUCCESS" if success else "❌ FAILED"
        print(f"   {file_path}: {status}")
        if success:
            successful_count += 1
    
    print(f"\n📈 Summary: {successful_count}/{len(results)} files executed successfully")
    
    if successful_count == len(results):
        print("🎉 All files executed successfully!")
    else:
        print("⚠️  Some files failed to execute. Please check the errors above.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n🛑 Execution interrupted by user (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error in master runner: {str(e)}")
        sys.exit(1)