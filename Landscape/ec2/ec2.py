import boto3
import json
import xlsxwriter
from datetime import datetime
import argparse
from botocore.exceptions import ClientError
import time
import os
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed

def get_ssm_patch_info(ssm_client, instance_id):
    try:
        current_patches = ssm_client.describe_instance_patches(
            InstanceId=instance_id,
            Filters=[{'Key': 'State', 'Values': ['Installed']}]
        )
        latest_patches = ssm_client.describe_available_patches()
        return {
            'current_patches': len(current_patches.get('Patches', [])),
            'latest_patches': len(latest_patches.get('Patches', []))
        }
    except Exception:
        return {'current_patches': 'N/A', 'latest_patches': 'N/A'}

def get_instance_os(ec2_client, instance):
    os_map = {
        'windows': 'Windows',
        'ubuntu': 'Ubuntu',
        'amazon linux': 'Amazon Linux',
        'red hat': 'RHEL',
        'rhel': 'RHEL',
        'suse': 'SUSE',
        'centos': 'CentOS',
        'debian': 'Debian'
    }

    try:
        if 'ImageId' not in instance:
            return 'N/A'
        response = ec2_client.describe_images(ImageIds=[instance['ImageId']])
        if not response['Images']:
            return 'N/A'
        image = response['Images'][0]
        description = image.get('Description', '').lower()
        platform = image.get('Platform', '').lower()
        name = image.get('Name', '').lower()
        if platform == 'windows':
            return 'Windows'
        for key, value in os_map.items():
            if key in description or key in name:
                return value
        return 'Linux'
    except Exception:
        return 'N/A'

def get_all_regions(ec2_client):
    try:
        return [region['RegionName'] for region in ec2_client.describe_regions()['Regions']]
    except Exception:
        return ['us-east-1']

def process_region(session, region, account_name):
    """Process a single region and return its instances"""
    instances_data = []
    try:
        regional_ec2 = session.client('ec2', region_name=region)
        regional_ssm = session.client('ssm', region_name=region)

        paginator = regional_ec2.get_paginator('describe_instances')
        for page in paginator.paginate():
            for reservation in page['Reservations']:
                for instance in reservation['Instances']:
                    state = instance['State']['Name']
                    if state == 'terminated':
                        continue
                    hostname = next(
                        (tag['Value'] for tag in instance.get('Tags', []) if tag['Key'] == 'Name'),
                        instance['InstanceId']
                    )
                    instances_data.append({
                        'Hostname': hostname,
                        'IsStopped': state == 'stopped',
                        'OS': get_instance_os(regional_ec2, instance),
                        'Region': region,
                        'InstanceType': instance['InstanceType'],
                        'CurrentPatches': get_ssm_patch_info(regional_ssm, instance['InstanceId'])['current_patches'],
                        'LatestPatches': get_ssm_patch_info(regional_ssm, instance['InstanceId'])['latest_patches']
                    })
    except Exception as e:
        print(f"Error processing region {region} in account {account_name}: {e}")
    
    return instances_data

def get_account_instances(account_name, access_key, secret_key):
    instances_data = []
    try:
        session = boto3.Session(
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key
        )
        ec2_client = session.client('ec2', region_name='us-east-1')
        regions = get_all_regions(ec2_client)

        print(f"  Processing {len(regions)} regions in parallel...")
        
        # Use ThreadPoolExecutor to process regions in parallel
        with ThreadPoolExecutor(max_workers=10) as executor:
            # Submit all region processing tasks
            future_to_region = {
                executor.submit(process_region, session, region, account_name): region 
                for region in regions
            }
            
            # Process results as they complete
            for future in as_completed(future_to_region):
                region = future_to_region[future]
                try:
                    region_instances = future.result()
                    instances_data.extend(region_instances)
                    print(f"    ✓ Completed region: {region} ({len(region_instances)} instances)")
                except Exception as e:
                    print(f"    ✗ Error in region {region}: {e}")
                
    except Exception as e:
        print(f"Error processing account {account_name}: {e}")
    return instances_data

def create_excel_report(account_name, instances, output_dir='reports'):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    filename = os.path.join(output_dir, f"{account_name.replace(' ', '_')}_landscape_{timestamp}.xlsx")
    workbook = xlsxwriter.Workbook(filename)

    cell_format = workbook.add_format({
        'font_name': 'Times New Roman',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    red_text = workbook.add_format({
        'font_name': 'Times New Roman',
        'font_color': 'red'
    })

    header_format = workbook.add_format({
        'bold': True,
        'font_name': 'Times New Roman',
        'font_size': 12,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    worksheet = workbook.add_worksheet('Landscape Summary')

    headers = [
        'S.No', 'Hostname', 'OS', 'Region',
        'Instance Type', 'Current Patch', 'Latest Patch'
    ]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    for row, instance in enumerate(instances, start=1):
        worksheet.write(row, 0, row, cell_format)

        # Rich formatting for Hostname with red [Stopped]
        if instance['IsStopped']:
            worksheet.write_rich_string(
                row, 1,
                cell_format, instance['Hostname'] + ' ',
                red_text, '[Stopped]',
                cell_format
            )
        else:
            worksheet.write(row, 1, instance['Hostname'], cell_format)

        worksheet.write(row, 2, instance['OS'], cell_format)
        worksheet.write(row, 3, instance['Region'], cell_format)
        worksheet.write(row, 4, instance['InstanceType'], cell_format)
        worksheet.write(row, 5, instance['CurrentPatches'], cell_format)
        worksheet.write(row, 6, instance['LatestPatches'], cell_format)

    col_widths = [8, 40, 15, 15, 15, 15, 15]
    for col, width in enumerate(col_widths):
        worksheet.set_column(col, col, width)

    worksheet.freeze_panes(1, 0)
    workbook.close()
    return filename

def load_accounts(filename='accounts.json'):
    try:
        with open(filename) as f:
            accounts = json.load(f)
            if not isinstance(accounts, list):
                print("Error: accounts.json should contain a list of accounts")
                return []
            valid_accounts = []
            for account in accounts:
                if not all(k in account for k in ['name', 'access_key', 'secret_key']):
                    print(f"Account {account.get('name', 'unknown')} missing required fields")
                    continue
                valid_accounts.append(account)
            return valid_accounts
    except Exception as e:
        print(f"Error loading accounts: {str(e)}")
        return []

def select_accounts(accounts):
    """Display numbered list of accounts and let user select which ones to process"""
    if not accounts:
        return []
    
    print("\n" + "="*60)
    print("AVAILABLE AWS ACCOUNTS")
    print("="*60)
    
    for i, account in enumerate(accounts, 1):
        print(f"{i}. {account['name']}")
    
    print("\n" + "="*60)
    print("SELECT ACCOUNTS TO PROCESS")
    print("="*60)
    print("Enter account numbers separated by commas (e.g., 1,3,5)")
    print("Enter 'all' to process all accounts")
    print("Enter 'none' to skip all accounts")
    
    while True:
        selection = input("\nYour selection: ").strip().lower()
        
        if selection == 'all':
            return accounts
        elif selection == 'none':
            return []
        elif selection:
            try:
                selected_indices = [int(idx.strip()) for idx in selection.split(',')]
                selected_accounts = []
                invalid_selections = []
                
                for idx in selected_indices:
                    if 1 <= idx <= len(accounts):
                        selected_accounts.append(accounts[idx-1])
                    else:
                        invalid_selections.append(idx)
                
                if invalid_selections:
                    print(f"Warning: Invalid account numbers ignored: {invalid_selections}")
                
                if selected_accounts:
                    print(f"\nSelected {len(selected_accounts)} account(s):")
                    for account in selected_accounts:
                        print(f"  - {account['name']}")
                    return selected_accounts
                else:
                    print("No valid accounts selected. Please try again.")
            except ValueError:
                print("Invalid input. Please enter numbers separated by commas.")
        else:
            print("No input received. Please enter your selection.")

def process_account(account, output_base_dir):
    """Process a single account and return results"""
    account_name = account['name']
    print(f"\nProcessing account: {account_name}")
    
    instances = get_account_instances(account_name, account['access_key'], account['secret_key'])
    
    if instances:
        report_path = create_excel_report(account_name, instances, output_base_dir)
        print(f"  ✅ Generated report: {report_path}")
        print(f"  ✅ Found {len(instances)} instances")
        return True, account_name, len(instances), report_path
    else:
        print("  ⚠️ No instances found for this account")
        return False, account_name, 0, None

def main():
    parser = argparse.ArgumentParser(
        description='AWS Landscape Information Collector',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument('--accounts-file', default='accounts.json', help='JSON file with account credentials')
    parser.add_argument('--output-dir', default='reports', help='Base directory for reports')
    args = parser.parse_args()

    print(f"Starting AWS Landscape collection at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Load accounts
    accounts = load_accounts(args.accounts_file)
    if not accounts:
        print("No accounts found to process")
        return
    
    # Let user select which accounts to process
    selected_accounts = select_accounts(accounts)
    if not selected_accounts:
        print("No accounts selected. Exiting.")
        return
    
    # Create timestamped reports directory
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    reports_dir = os.path.join(args.output_dir, f"reports_{timestamp}")
    os.makedirs(reports_dir, exist_ok=True)
    
    print(f"\n" + "="*60)
    print(f"PROCESSING {len(selected_accounts)} SELECTED ACCOUNT(S)")
    print(f"Reports will be saved to: {reports_dir}")
    print("="*60)
    
    # Process accounts in parallel
    print(f"\nProcessing accounts with {min(5, len(selected_accounts))} workers...")
    
    results = []
    with ThreadPoolExecutor(max_workers=min(5, len(selected_accounts))) as executor:
        # Submit all account processing tasks
        future_to_account = {
            executor.submit(process_account, account, reports_dir): account['name']
            for account in selected_accounts
        }
        
        # Process results as they complete
        for future in as_completed(future_to_account):
            account_name = future_to_account[future]
            try:
                success, processed_name, instance_count, report_path = future.result()
                results.append({
                    'account': processed_name,
                    'success': success,
                    'instance_count': instance_count,
                    'report_path': report_path
                })
            except Exception as e:
                print(f"Error processing account {account_name}: {e}")
                results.append({
                    'account': account_name,
                    'success': False,
                    'instance_count': 0,
                    'report_path': None,
                    'error': str(e)
                })
    
    # Print summary
    print(f"\n" + "="*60)
    print("PROCESSING COMPLETE - SUMMARY")
    print("="*60)
    
    total_instances = 0
    successful_accounts = 0
    
    for result in results:
        if result['success']:
            print(f"✅ {result['account']}: {result['instance_count']} instances")
            total_instances += result['instance_count']
            successful_accounts += 1
        else:
            print(f"❌ {result['account']}: Failed to process")
    
    print(f"\n📊 SUMMARY:")
    print(f"   Processed accounts: {successful_accounts}/{len(selected_accounts)}")
    print(f"   Total instances found: {total_instances}")
    print(f"   Reports saved to: {reports_dir}")
    print(f"\n✅ Completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == '__main__':
    main()
