// Tells Tailwind which file to scan for style classes.
// Rebuild tailwind.css after changing classes in index.html:
//   npx tailwindcss@3 -i tailwind.input.css -o tailwind.css --minify
module.exports = {
  content: ['./index.html'],
  theme: { extend: {} },
  plugins: [],
};
