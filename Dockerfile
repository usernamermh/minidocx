FROM harbor.unisound.ai/unisound/slime:v1

ENV DEBIAN_FRONTEND=noninteractive

# Install SSH server and basic user-management tools.
RUN apt-get update \
    && apt-get install -y --no-install-recommends openssh-server sudo \
    && rm -rf /var/lib/apt/lists/*

# Create required groups with fixed gids when they do not already exist.
RUN getent group 3000 >/dev/null || groupadd -g 3000 unisound \
    && getent group 3001 >/dev/null || groupadd -g 3001 docker \
    && getent group 3002 >/dev/null || groupadd -g 3002 nlp \
    && getent group 2069 >/dev/null || groupadd -g 2069 renminhui

# Create user and attach supplementary groups.
RUN id -u renminhui >/dev/null 2>&1 || useradd -m -u 2069 -g 2069 -G 3000,3001,3002 -s /bin/bash renminhui \
    && echo 'renminhui:Aa!51801739462' | chpasswd \
    && usermod -aG sudo renminhui

# Enable password login for SSH.
RUN mkdir -p /var/run/sshd \
    && sed -i 's/^#\\?PasswordAuthentication .*/PasswordAuthentication yes/' /etc/ssh/sshd_config \
    && sed -i 's/^#\\?PermitRootLogin .*/PermitRootLogin no/' /etc/ssh/sshd_config \
    && sed -i 's/^#\\?UsePAM .*/UsePAM yes/' /etc/ssh/sshd_config \
    && printf '\nAllowUsers renminhui\n' >> /etc/ssh/sshd_config

EXPOSE 22

CMD ["/usr/sbin/sshd", "-D"]
