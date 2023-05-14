from google.cloud import storage
import subprocess

# Create a client using the service account credentials
client = storage.Client()

# Get a reference to the bucket and the blob (packages.tar.gz file)
bucket = client.get_bucket('parecengine')
blob = bucket.blob('vendor/packages.tar.gz')

# Download the packages.tar.gz file to a local directory
blob.download_to_filename('/app/packages.tar.gz')

# Extract the packages.tar.gz file
subprocess.run(['tar', '-xzf', '/app/packages.tar.gz', '-C', '/app'])

# Install the packages using pip
subprocess.run(['pip', 'install', '--no-index', '--find-links', '/app/vendor/packages', '-r', 'requirements.txt'])