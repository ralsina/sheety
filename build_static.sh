#!/bin/bash
set -e

docker run --rm --privileged \
  multiarch/qemu-user-static \
  --reset -p yes

# Build for AMD64
docker build . -f Dockerfile.static -t sheety-builder
docker run -i --rm -v "$PWD":/app --user="$UID" sheety-builder /bin/sh -c "cd /app && rm -rf lib shard.lock && shards build --release --without-development --static"
mv bin/sheety bin/sheety-static-linux-amd64

# Build for ARM64
docker build . -f Dockerfile.static --platform linux/arm64 -t sheety-builder-arm64
docker run -i --rm -v "$PWD":/app --platform linux/arm64 --user="$UID" sheety-builder-arm64 /bin/sh -c "cd /app && rm -rf lib shard.lock && shards build --release --without-development --static"
mv bin/sheety bin/sheety-static-linux-arm64
