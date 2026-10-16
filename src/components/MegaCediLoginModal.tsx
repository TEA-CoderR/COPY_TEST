import React from 'react';
import { PlatformAuthModal } from './PlatformAuthModal';

interface MegaCediLoginModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSessionUpdated?: () => void;
}

export const MegaCediLoginModal: React.FC<MegaCediLoginModalProps> = ({
  isOpen,
  onClose,
  onSessionUpdated,
}) => {
  return (
    <PlatformAuthModal
      isOpen={isOpen}
      onClose={onClose}
      initialPlatform="megacedi"
      onSessionUpdated={onSessionUpdated}
    />
  );
};
