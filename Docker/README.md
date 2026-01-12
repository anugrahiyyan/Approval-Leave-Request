############################################
# If you make an update in this repository #
# Please, do this step becarefully         #
############################################

###### Stop and delete old docker run ######
sudo docker stop leave-app && sudo docker rm leave-app

###### Update the new change on leave/privacy/terms page ######
sudo docker build -t leave-app-system .

###### Run newest docker update ######
sudo docker run -d -p 8553:80 --name leave-app leave-app-system

###### How to publish maintenance site ######
Maintenance Active = ./down.sh
Maintenance Done   = ./up.sh
