import UploadExcel from './components/UploadExcel';
import { AppBar, Toolbar, Typography, Box, CssBaseline, ThemeProvider, createTheme } from '@mui/material';
import logo from './assets/logo.png';

// 1. Cambiar a un tema claro, manteniendo el color primario
const theme = createTheme({
  palette: {
    mode: 'light', // Cambiado a modo claro
    primary: {
      main: '#FFA000', // Tono de amarillo más oscuro
    },
    secondary: {
      main: '#6c757d', // Gris secundario estándar
    },
    background: {
      default: '#f4f6f8', // Fondo general gris muy claro
      paper: '#ffffff',   // Fondo de componentes blanco
    },
    text: {
      primary: '#212529', // Texto principal oscuro
      secondary: '#6c757d', // Texto secundario gris
    },
  },
  typography: {
    fontFamily: 'Roboto, Arial, sans-serif',
    h5: {
      fontWeight: 300,
    },
  },
});

function App() {
  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Box sx={{ flexGrow: 1 }}>
        {/* 2. Anular el tema para que el AppBar siga siendo oscuro */}
        <AppBar 
          position="static"
          sx={{ 
            backgroundColor: '#1E1E1E', // Fondo oscuro específico para el header
            boxShadow: '0 2px 4px -1px rgba(0,0,0,0.2)' 
          }}
        >
          <Toolbar>
            <Box
              component="img"
              sx={{
                height: 40,
                mr: 2,
              }}
              alt="Logo"
              src={logo}
            />
            <Typography variant="h5" component="div" sx={{ flexGrow: 1, color: '#FFFFFF' /* Texto blanco para el header */ }}>
              Cargar Beneficiarios Generac
            </Typography>
          </Toolbar>
        </AppBar>
        <Box component="main" sx={{ p: { xs: 1, sm: 2, md: 3 } }}>
          <UploadExcel />
        </Box>
      </Box>
    </ThemeProvider>
  );
}

export default App;
