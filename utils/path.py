import os
from dotenv import load_dotenv
load_dotenv()

def createEnvPath(env_path, last = None):
    if os.getenv('MODE') == 'production':
        if last:
            return os.path.join(os.getcwd(), 'dist', os.getenv(env_path), last)
        return os.path.join(os.getcwd(), 'dist', os.getenv(env_path))
    else:
        if last:
            return os.path.join(os.getcwd(), os.getenv(env_path), last)
    return os.path.join(os.getcwd(), os.getenv(env_path))
