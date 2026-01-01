import boto3
import csv
from datetime import datetime

OUTPUT_FILE = f"unused_ebs_volumes_{datetime.now().strftime('%Y%m%d')}.csv"


def get_all_regions():
    ec2 = boto3.client("ec2")
    regions = ec2.describe_regions()["Regions"]
    return [region["RegionName"] for region in regions]


def get_unused_volumes(region):
    ec2 = boto3.client("ec2", region_name=region)
    paginator = ec2.get_paginator("describe_volumes")

    unused_volumes = []

    for page in paginator.paginate(
        Filters=[{"Name": "status", "Values": ["available"]}]
    ):
        for volume in page["Volumes"]:
            volume_name = "N/A"
            for tag in volume.get("Tags", []):
                if tag["Key"].lower() == "name":
                    volume_name = tag["Value"]

            unused_volumes.append({
                "Region": region,
                "Volume ID": volume["VolumeId"],
                "Name": volume_name,
                "Size (GB)": volume["Size"],
                "Volume Type": volume["VolumeType"],
                "State": volume["State"],
                "Created On": volume["CreateTime"].strftime("%Y-%m-%d %H:%M:%S")
            })

    return unused_volumes


def main():
    print("Starting unused EBS volume audit...\n")

    all_regions = get_all_regions()
    all_unused_volumes = []

    for region in all_regions:
        print(f"Scanning region: {region}")
        volumes = get_unused_volumes(region)
        all_unused_volumes.extend(volumes)

    if not all_unused_volumes:
        print("\nNo unused EBS volumes found.")
        return

    with open(OUTPUT_FILE, "w", newline="") as csvfile:
        fieldnames = all_unused_volumes[0].keys()
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        writer.writerows(all_unused_volumes)

    print(f"\nAudit completed successfully.")
    print(f"Report generated: {OUTPUT_FILE}")
    print(f"Total unused volumes found: {len(all_unused_volumes)}")


if __name__ == "__main__":
    main()
