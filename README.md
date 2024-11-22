# 🖼️ Slide Finder

**Slide Finder** is a simple utility to streamline PowerPoint presentations by comparing a master slide deck with custom slides. It helps identify and manage variations between master and custom slides, ensuring presentation consistency and clarity.

Use Slide Finder to identify which slides in your presentation need adjustments after changes have been made to the master slides.

---

## ✨ Features

- **Master-Your-Slides Comparison**: Compare a master slide deck with your customized slides.   
- **Drag-and-Drop Support**: Simply drag your slides onto the executable to bypass configuration.  
- **Easy Setup**: Configure paths to your PowerPoint files in the `app.config` file.

---

## 🛠️ Configuration and Usage

To configure and run the tool, follow these steps:

### 1. Update `app.config`
The configuration file `app.config` defines the paths to your PowerPoint files. Open it in a text editor and adjust the following settings under `<appSettings>`:

```xml
<appSettings>
  <!-- Fully qualified path to the master slide deck -->
  <add key="masterSlides" value="C:\Users\xyz\Downloads\Kickoff-Workshop\MASTER slide deck.pptx" />

  <!-- Fully qualified path to your custom slides -->
  <add key="yourSlides" value="C:\Users\xyz\Downloads\Kickoff-Workshop\product-overview.pptx" />
</appSettings>
```

#### Key Notes:
- **`masterSlides`**: Path to your master slide deck file (`.pptx`).  
- **`yourSlides`**: Path to your customized slides file (`.pptx`).  

> **Drag-and-Drop Support**: If you drag and drop your slides onto the executable, the `yourSlides` value will be ignored. Only the master slides (`masterSlides`) will be used for comparison.

---

### 2. Requirements

- **Supported Runtime**: `.NET Runtime.  
- **PowerPoint Files**: Ensure your `.pptx` files are valid and accessible at the specified paths.

---

### 3. Run the Tool

- **Default Configuration**: Double-click the `SlideFinder.exe` to run using the `app.config` settings.  
- **Drag-and-Drop**: Drag your slide deck (not the master slides!) onto the executable for quick comparison.

---

## 🧩 How It Works

1. Compares each slide in the specified custom deck (`yourSlides`) against the master deck (`masterSlides`).
2. Flags differences to help streamline adjustments and maintain presentation consistency.

---

## 🤝 Contributions

Contributions are welcome! Feel free to open issues or submit pull requests via the [Issues tab](https://github.com/nenning/Slide-Finder/issues).

---

## 📜 License

This project is distributed under the [MIT License](LICENSE).
