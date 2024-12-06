#!/bin/bash
# Domotz script to check connection to the Domotz Cloud and default dns for Linux
# it uses bash internals /dev/tcp utils and ping.
# Domotz startup script

TIMEOUT="5"
DNSADDR=$(cat /etc/resolv.conf |grep -i '^nameserver'|head -n1|cut -d ' ' -f2)
TCPCHECK="timeout ${TIMEOUT} bash -c"

DECHO="echo.domotz.com"
PORTAL="portal.domotz.com"
PORTALP="443"
APIEU="api-eu-west-1-cell-1.domotz.com"
PAPIEU="443"
RBTEU="messaging-eu-west-1-cell-1.domotz.com"
PRBTEU="5671"
APIUS="api-us-east-1-cell-1.domotz.com"
PAPIUS="443"
RBTUS="messaging-us-east-1-cell-1.domotz.com"
PRBTUS="5671"

echo "+------------------------------------------------+"
echo "|  ___                             _             |"
echo "| (  _'\                          ( )_           |"
echo "| | | ) |   _     ___ ___     _   | ,_) ____     |"
echo "| | | | ) /'_'\ /' _ ' _ '\ /'_'\ | |  (_  ,)    |"
echo "| | |_) |( (_) )| ( ) ( ) |( (_) )| |_  /'/_     |"
echo "| (____/''\___/'(_) (_) (_)'\___/''\__)(____)    |"
echo "| ---------------------------------------------- |"
echo "| The IT Monitoring and Management Solition      |"
echo "+------------------------------------------------+"
echo "=================================================="

echo "This script checks that the connection to the Domotz Cloud is reliable"
echo ""

echo "In which area is your Domotz Agent located?"
PS3='Please enter 1,2,3 or 4:'
options=("USA" "Europe" "APAC" "Quit")
select opt in "${options[@]}"
do
    case $opt in
        "USA")
            ZONE="us"
            break
            ;;
        "Europe")
            ZONE="eu"
            break
            ;;
        "APAC")
            ZONE="apac"
            break
            ;;
        "Quit")
            exit
            ;;
        *) echo "invalid option $REPLY";;
    esac
done

clear

echo ""
if [ "${DNSADDR}" = "8.8.8.8" ]; then

    echo "-> DNS settings ok!"
else
    echo "Can you please make sure to use at least one DNS server option as a public one like the Google DNS server (8.8.8.8 or 8.8.4.4)?"
fi

if ping -c 1 -W 1 "${DECHO}" > /dev/null; then
 S_DECHO="ok"
fi
if [ "${S_DECHO}" = "ok" ]; then
    echo "-> ${DECHO} ok!"
else
    echo "-> ${DECHO} KO!! NOT REACHABLE"
fi

echo ""
${TCPCHECK} "</dev/tcp/${PORTAL}/${PORTALP}" && echo "${PORTAL}/${PORTALP} ok!"|| echo "${PORTAL}/${PORTALP} KO!! NOT REACHABLE"
echo ""

if [ "${ZONE}" = "eu" ]; then
    echo "-> TESTING EU SERVERS"
    ${TCPCHECK} "</dev/tcp/${APIEU}/${PAPIEU}" && echo "${APIEU}:${PAPIEU} ok!" || echo "${APIEU}:${PAPIEU} KO!! NOT REACHABLE"
    ${TCPCHECK} "</dev/tcp/${RBTEU}/${PRBTEU}" && echo "${RBTEU}:${PRBTEU} ok!" || echo "${RBTEU}:${PRBTEU} KO!! NOT REACHABLE"
fi 

if [ "${ZONE}" = "apac" ]; then
    echo "-> TESTING APAC SERVERS"
    ${TCPCHECK} "</dev/tcp/${APIEU}/${PAPIEU}" && echo "${APIEU}:${PAPIEU} ok!" || echo "${APIEU}:${PAPIEU} KO!! NOT REACHABLE"
    ${TCPCHECK} "</dev/tcp/${RBTEU}/${PRBTEU}" && echo "${RBTEU}:${PRBTEU} ok!" || echo "${RBTEU}:${PRBTEU} KO!! NOT REACHABLE"
fi 


if [ "${ZONE}" = "us" ]; then
    echo "-> TESTING US SERVERS"
    ${TCPCHECK} "</dev/tcp/${APIUS}/${PAPIUS}" && echo "${APIUS}:${PAPIUS} ok!" || echo "${APIUS}:${PAPIUS} KO!! NOT REACHABLE"
    ${TCPCHECK} "</dev/tcp/${RBTUS}/${PRBTUS}" && echo "${RBTUS}:${PRBTUS} ok!" || echo "${RBTUS}:${PRBTUS} KO!! NOT REACHABLE"
fi 

echo ""
echo "N.B. To remotely connect to your devices  please make sure that the following host/port-range is allowed on your firewall:"
if [ "${ZONE}" = "eu" ]; then
    echo "sshg.domotz.co (range: 32700 - 57699 TCP)"
fi
if [ "${ZONE}" = "apac" ]; then
    echo "ap-southeast-2-sshg.domotz.co(range: 32700 - 57699 TCP)"
fi
if [ "${ZONE}" = "us" ]; then
    echo "us-east-1-sshg.domotz.co, us-east-1-02-sshg.domotz.co and us-west-2-sshg.domotz.co (range: 32700 - 57699 TCP)"
fi