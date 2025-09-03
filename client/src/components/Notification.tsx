import React from 'react';
import { useNotification } from '../hooks/useNotification';

interface NotificationProps {
  message: string;
  type: 'info' | 'success' | 'error' | 'warning';
  onClose: () => void;
}

const NotificationItem: React.FC<NotificationProps> = ({ message, type, onClose }) => {
  const getBackgroundColor = () => {
    switch (type) {
      case 'success':
        return '#4CAF50';
      case 'error':
        return '#f44336';
      case 'warning':
        return '#ff9800';
      default:
        return '#217346';
    }
  };

  return (
    <div
      style={{
        position: 'fixed',
        top: '20px',
        right: '20px',
        backgroundColor: getBackgroundColor(),
        color: 'white',
        padding: '12px 20px',
        borderRadius: '6px',
        boxShadow: '0 4px 8px rgba(0,0,0,0.2)',
        zIndex: 1001,
        fontSize: '14px',
        fontWeight: 500,
        opacity: 1,
        transition: 'opacity 0.3s ease',
        cursor: 'pointer'
      }}
      onClick={onClose}
    >
      {message}
    </div>
  );
};

export const NotificationContainer: React.FC = () => {
  const { notifications, removeNotification } = useNotification();

  return (
    <>
      {notifications.map((notification, index) => (
        <div
          key={notification.id}
          style={{
            position: 'fixed',
            top: `${20 + index * 60}px`,
            right: '20px',
            zIndex: 1001
          }}
        >
          <NotificationItem
            message={notification.message}
            type={notification.type || 'info'}
            onClose={() => removeNotification(notification.id)}
          />
        </div>
      ))}
    </>
  );
};

export default NotificationContainer;