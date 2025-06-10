#!/usr/bin/env python3
"""Google Cloud Function deployment script."""

import json
import os
import sys
import subprocess
import tempfile
import zipfile
from pathlib import Path
import click

def run_command(command, description="", check=True):
    """Run a shell command."""
    print(f"Running: {command}")
    if description:
        print(f"  {description}")
    
    try:
        result = subprocess.run(command, shell=True, check=check, capture_output=True, text=True)
        if result.stdout:
            print(f"  Output: {result.stdout.strip()}")
        return result
    except subprocess.CalledProcessError as e:
        print(f"  Error: {e}")
        if e.stderr:
            print(f"  Stderr: {e.stderr.strip()}")
        if check:
            sys.exit(1)
        return e

def check_gcloud_installed():
    """Check if gcloud CLI is installed."""
    result = run_command("gcloud --version", "Checking gcloud installation", check=False)
    if result.returncode != 0:
        print("Error: gcloud CLI is not installed")
        print("Please install it from: https://cloud.google.com/sdk/docs/install")
        sys.exit(1)
    print("✓ gcloud CLI is installed")

def check_gcloud_authenticated():
    """Check if user is authenticated with gcloud."""
    result = run_command("gcloud auth list --filter=status:ACTIVE --format='value(account)'", 
                        "Checking gcloud authentication", check=False)
    if not result.stdout.strip():
        print("Error: Not authenticated with gcloud")
        print("Please run: gcloud auth login")
        sys.exit(1)
    print(f"✓ Authenticated as: {result.stdout.strip()}")

def validate_project_config(project_id, region, function_name):
    """Validate project configuration."""
    if not project_id:
        print("Error: PROJECT_ID is required")
        sys.exit(1)
    
    if not region:
        print("Error: REGION is required")
        sys.exit(1)
    
    if not function_name:
        print("Error: FUNCTION_NAME is required")
        sys.exit(1)
    
    print(f"✓ Project ID: {project_id}")
    print(f"✓ Region: {region}")
    print(f"✓ Function Name: {function_name}")

def enable_required_apis(project_id):
    """Enable required Google Cloud APIs."""
    required_apis = [
        'cloudfunctions.googleapis.com',
        'cloudbuild.googleapis.com',
        'storage.googleapis.com',
        'logging.googleapis.com'
    ]
    
    print("Enabling required APIs...")
    for api in required_apis:
        run_command(f"gcloud services enable {api} --project={project_id}",
                   f"Enabling {api}")
    
    print("✓ Required APIs enabled")

def prepare_deployment_package():
    """Prepare deployment package."""
    print("Preparing deployment package...")
    
    # Create temporary directory for deployment files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Copy source files
        src_dest = temp_path / 'src'
        src_dest.mkdir()
        
        project_root = Path(__file__).parent.parent
        src_path = project_root / 'src'
        
        # Copy Python files
        for py_file in src_path.rglob('*.py'):
            rel_path = py_file.relative_to(src_path)
            dest_file = src_dest / rel_path
            dest_file.parent.mkdir(parents=True, exist_ok=True)
            dest_file.write_text(py_file.read_text())
        
        # Copy requirements.txt
        requirements_src = project_root / 'requirements.txt'
        requirements_dest = temp_path / 'requirements.txt'
        requirements_dest.write_text(requirements_src.read_text())
        
        # Copy config files
        config_src = project_root / 'config'
        config_dest = temp_path / 'config'
        if config_src.exists():
            config_dest.mkdir()
            for config_file in config_src.glob('*.json'):
                (config_dest / config_file.name).write_text(config_file.read_text())
        
        # Create main.py for Cloud Functions
        main_py_content = """
import functions_framework
from src.main import excel_pptx_merger

# Cloud Function entry point
@functions_framework.http
def main(request):
    return excel_pptx_merger(request)
"""
        (temp_path / 'main.py').write_text(main_py_content.strip())
        
        # Create deployment zip
        zip_path = project_root / 'deployment.zip'
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in temp_path.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_path)
                    zipf.write(file_path, arcname)
        
        print(f"✓ Deployment package created: {zip_path}")
        return zip_path

def deploy_function(project_id, region, function_name, zip_path, env_vars):
    """Deploy function to Google Cloud."""
    print(f"Deploying function: {function_name}")
    
    # Build deployment command
    command = [
        "gcloud", "functions", "deploy", function_name,
        f"--source={zip_path}",
        "--entry-point=main",
        "--runtime=python311",
        "--trigger=http",
        "--allow-unauthenticated",
        f"--region={region}",
        f"--project={project_id}",
        "--memory=1024MB",
        "--timeout=540s"
    ]
    
    # Add environment variables
    if env_vars:
        env_vars_str = ",".join([f"{k}={v}" for k, v in env_vars.items()])
        command.extend(["--set-env-vars", env_vars_str])
    
    # Execute deployment
    result = run_command(" ".join(command), "Deploying Cloud Function")
    
    if result.returncode == 0:
        print("✓ Function deployed successfully")
        
        # Get function URL
        url_result = run_command(
            f"gcloud functions describe {function_name} --region={region} --project={project_id} --format='value(httpsTrigger.url)'",
            "Getting function URL"
        )
        
        if url_result.stdout.strip():
            print(f"✓ Function URL: {url_result.stdout.strip()}")
    else:
        print("✗ Function deployment failed")
        sys.exit(1)

def setup_cloud_storage(project_id, bucket_name):
    """Set up Cloud Storage bucket if specified."""
    if not bucket_name:
        return
    
    print(f"Setting up Cloud Storage bucket: {bucket_name}")
    
    # Check if bucket exists
    result = run_command(
        f"gsutil ls gs://{bucket_name}/",
        "Checking if bucket exists",
        check=False
    )
    
    if result.returncode != 0:
        # Create bucket
        run_command(
            f"gsutil mb -p {project_id} gs://{bucket_name}",
            "Creating Cloud Storage bucket"
        )
        print(f"✓ Created bucket: gs://{bucket_name}")
    else:
        print(f"✓ Bucket already exists: gs://{bucket_name}")

@click.command()
@click.option('--project-id', envvar='GOOGLE_CLOUD_PROJECT', required=True,
              help='Google Cloud Project ID')
@click.option('--region', default='us-central1', 
              help='Google Cloud region for deployment')
@click.option('--function-name', default='excel-pptx-merger',
              help='Cloud Function name')
@click.option('--bucket-name', envvar='GOOGLE_CLOUD_BUCKET',
              help='Cloud Storage bucket name (optional)')
@click.option('--env', default='production',
              help='Environment (development, testing, production)')
@click.option('--api-key', envvar='API_KEY',
              help='API key for authentication')
@click.option('--dry-run', is_flag=True,
              help='Show what would be deployed without actually deploying')
def deploy(project_id, region, function_name, bucket_name, env, api_key, dry_run):
    """Deploy Excel to PowerPoint Merger to Google Cloud Functions."""
    
    print("Excel to PowerPoint Merger - Google Cloud Deployment")
    print("=" * 60)
    
    if dry_run:
        print("DRY RUN MODE - No actual deployment will occur")
        print()
    
    # Validation
    validate_project_config(project_id, region, function_name)
    
    if not dry_run:
        check_gcloud_installed()
        check_gcloud_authenticated()
    
    # Environment variables
    env_vars = {
        'ENVIRONMENT': env,
        'DEVELOPMENT_MODE': 'false',
        'LOG_LEVEL': 'INFO',
        'CLEANUP_TEMP_FILES': 'true',
        'TEMP_FILE_RETENTION_SECONDS': '300'
    }
    
    if api_key:
        env_vars['API_KEY'] = api_key
    
    if bucket_name:
        env_vars['GOOGLE_CLOUD_BUCKET'] = bucket_name
    
    print(f"Environment variables: {list(env_vars.keys())}")
    
    if dry_run:
        print("\nDry run complete - deployment configuration is valid")
        return
    
    try:
        # Enable APIs
        enable_required_apis(project_id)
        
        # Set up Cloud Storage if needed
        if bucket_name:
            setup_cloud_storage(project_id, bucket_name)
        
        # Prepare deployment package
        zip_path = prepare_deployment_package()
        
        # Deploy function
        deploy_function(project_id, region, function_name, zip_path, env_vars)
        
        # Cleanup
        if zip_path.exists():
            zip_path.unlink()
            print("✓ Cleaned up deployment package")
        
        print("\n" + "="*60)
        print("🎉 Deployment completed successfully!")
        print("="*60)
        print(f"\nFunction: {function_name}")
        print(f"Project: {project_id}")
        print(f"Region: {region}")
        print(f"Environment: {env}")
        
        if bucket_name:
            print(f"Storage bucket: gs://{bucket_name}")
        
        print("\nNext steps:")
        print("1. Test the deployment with the function URL")
        print("2. Set up monitoring and alerts in Google Cloud Console")
        print("3. Configure custom domain if needed")
        print("4. Set up CI/CD pipeline for automated deployments")
        
    except Exception as e:
        print(f"\n✗ Deployment failed: {e}")
        sys.exit(1)

if __name__ == '__main__':
    deploy()