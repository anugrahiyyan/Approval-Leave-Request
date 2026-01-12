#!/bin/bash

echo "🔧 Enabling maintenance mode..."
sudo docker exec -it leave-app touch /var/www/html/down/down.flag
echo "✅ Maintenance mode enabled. All requests will now show maintenance page."
