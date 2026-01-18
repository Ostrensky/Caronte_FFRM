# --- FILE: app/shared_memory.py ---
import pandas as pd
import tempfile
import os
import logging
import uuid

# Global dictionary to hold DataFrames by ID
_SHARED_DF_CACHE = {}

def share_dataframe(df):
    """Saves a DataFrame to a temporary location and returns its unique ID."""
    global _SHARED_DF_CACHE
    df_id = str(uuid.uuid4())
    
    # In this simple model, we just hold the DF in memory by ID.
    # In a real-world scenario, you might save it to a temp file and return the path,
    # but for multiprocessing, holding the reference in a shared dictionary in the parent
    # often works cleaner than relying on temp disk I/O for the child process to read.
    # However, since the child process is a *copy* of the parent's environment, 
    # and to ensure data integrity during pickling, we will rely on temporary disk I/O.
    
    try:
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, f'shared_df_{df_id}.pkl')
        
        # Save the DataFrame to disk for the child to read
        df.to_pickle(file_path)
        
        logging.info(f"Shared DF saved to: {file_path}")
        return {'id': df_id, 'path': file_path}
    except Exception as e:
        logging.error(f"Failed to save DataFrame for sharing: {e}")
        raise

def retrieve_dataframe(shared_info):
    """Reads the DataFrame from the temporary location and cleans up."""
    file_path = shared_info['path']
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Shared DataFrame file not found at: {file_path}")
        
    try:
        # Load the DataFrame
        df = pd.read_pickle(file_path)
        return df
    finally:
        # Crucial: Delete the temporary file immediately after retrieval
        try:
            os.remove(file_path)
            logging.info(f"Cleaned up shared DF file: {file_path}")
        except Exception as e:
            logging.warning(f"Failed to clean up shared DF file: {e}")