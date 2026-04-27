import os
import glob
import shutil

def cleanup_old_sessions(directories, current_session):
    """
    Purges ghost artifact files from previous instances locking them to maintain UI performance.
    """
    print(f"Running cleanup. Protecting current session: {current_session}")
    for directory in directories:
        if not os.path.exists(directory):
            continue
        for file in glob.glob(os.path.join(directory, "*")):
            filename = os.path.basename(file)
            if current_session not in filename:
                try:
                    # Ignore .gitignores if present
                    if not filename.startswith("."): 
                        if os.path.isdir(file):
                            shutil.rmtree(file)
                        else:
                            os.remove(file)
                        print(f"Purged old ghost artifact: {filename}")
                except Exception as e:
                    print(f"Warning: Could not delete {filename}. {e}")
