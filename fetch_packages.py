import subprocess

# Install each package from the cloud storage bucket
subprocess.check_call(['gsutil', '-m', 'cp', '-r', 'gs://parecengine/vendor', 'vendor'])
subprocess.check_call(['pip', 'install', '--no-index', '--find-links', 'vendor', '-r', 'requirements.txt'])